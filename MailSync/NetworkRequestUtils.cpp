//
//  NetworkRequestUtils.cpp
//  MailSync
//
//  Created by Ben Gotow on 8/6/17.
//  Copyright © 2017 Foundry 376. All rights reserved.
//
//  Use of this file is subject to the terms and conditions defined
//  in 'LICENSE.md', which is part of the Mailspring-Sync package.
//

#include "Identity.hpp"
#include "NetworkRequestUtils.hpp"
#include "SyncException.hpp"
#include "Account.hpp"
#include "MailUtils.hpp"

#include <sys/stat.h>
#include <vector>
#include <cstdlib>
#include <cstring>

// Structure to hold curl request data that needs cleanup
// Stored in CURLOPT_PRIVATE to enable proper cleanup of headers
struct CurlRequestData {
    struct curl_slist *headers;

    CurlRequestData() : headers(nullptr) {}
};

// Helper to properly cleanup a curl handle including headers stored in private data
void CleanupCurlRequest(CURL * curl_handle) {
    if (curl_handle == nullptr) return;

    CurlRequestData *data = nullptr;
    curl_easy_getinfo(curl_handle, CURLINFO_PRIVATE, &data);
    if (data != nullptr) {
        if (data->headers != nullptr) {
            curl_slist_free_all(data->headers);
        }
        delete data;
    }
    curl_easy_cleanup(curl_handle);
}

size_t _onAppendToString(void *contents, size_t length, size_t nmemb, void *userp) {
    string * buffer = (string *)userp;
    size_t real_size = length * nmemb;
    
    size_t oldLength = buffer->size();
    size_t newLength = oldLength + real_size;
    
    buffer->resize(newLength);
    std::copy((char*)contents, (char*)contents+real_size, buffer->begin() + oldLength);
    
    return real_size;
}

const json MakeOAuthRefreshRequest(string provider, string clientId, string refreshToken, string scope) {
    CURL * curl_handle = curl_easy_init();
    const char * url =
          provider == "gmail" ? "https://www.googleapis.com/oauth2/v4/token"
        : provider == "office365" ? "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        : provider == "outlook" ? "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        : "";
    curl_easy_setopt(curl_handle, CURLOPT_URL, url);
    curl_easy_setopt(curl_handle, CURLOPT_CONNECTTIMEOUT, 20);

    char * c = curl_easy_escape(curl_handle, clientId.c_str(), 0);
    char * r = curl_easy_escape(curl_handle, refreshToken.c_str(), 0);
    string payload = "grant_type=refresh_token&client_id=" + string(c) + "&refresh_token=" + string(r);
    curl_free(c);
    curl_free(r);

    if (provider == "office365" || provider == "outlook") {
        // workaround the fact that Microsoft's OAUTH flow allows you to authorize many scopes, but you
        // have to get a separate token for outlook (email + IMAP) and contacts / calendar / Microsoft Graph APIs
        // separately. The same refresh token will give you access tokens, but the access tokens are different.
        if (scope == "") {
            scope = "https://outlook.office.com/IMAP.AccessAsUser.All https://outlook.office.com/SMTP.Send";
        }
        char * s = curl_easy_escape(curl_handle, scope.c_str(), 0);
        payload += "&scope=" + string(s);
        curl_free(s);
    }

    string gmailClientId = MailUtils::getEnvUTF8("GMAIL_CLIENT_ID");
    string gmailClientSecret = MailUtils::getEnvUTF8("GMAIL_CLIENT_SECRET");
    if (provider == "gmail" && clientId == gmailClientId) {
        // per https://stackoverflow.com/questions/59416326/safely-distribute-oauth-2-0-client-secret-in-desktop-applications-in-python,
        // we really do need to embed this in the application and it's more an extension of the Client ID than a proper Client Secret.
        // For a full explanation, see onboarding-helpers.ts in Mailspring. Please don't re-use this client id + secret in derivative
        // works or other products.
        payload += "&client_secret=" + gmailClientSecret;
    }

    // Store headers in CurlRequestData for proper cleanup
    CurlRequestData *requestData = new CurlRequestData();
    requestData->headers = curl_slist_append(requestData->headers, "Accept: application/json");
    requestData->headers = curl_slist_append(requestData->headers, "Content-Type: application/x-www-form-urlencoded");

    curl_easy_setopt(curl_handle, CURLOPT_CUSTOMREQUEST, "POST");
    curl_easy_setopt(curl_handle, CURLOPT_HTTPHEADER, requestData->headers);
    curl_easy_setopt(curl_handle, CURLOPT_POSTFIELDS, payload.c_str());
    curl_easy_setopt(curl_handle, CURLOPT_PRIVATE, requestData);

    return PerformJSONRequest(curl_handle);
}

CURL * CreateJSONRequest(string url, string method, string authorization, const char * payloadChars) {
    CURL * curl_handle = curl_easy_init();
    curl_easy_setopt(curl_handle, CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl_handle, CURLOPT_CONNECTTIMEOUT, 20);

    // Store headers in CurlRequestData for proper cleanup
    CurlRequestData *requestData = new CurlRequestData();

    requestData->headers = curl_slist_append(requestData->headers, "Accept: application/json");

    if (authorization != "") {
        requestData->headers = curl_slist_append(requestData->headers, ("Authorization: " + authorization).c_str());
    }
    if (payloadChars != nullptr && strlen(payloadChars) > 0) {
        requestData->headers = curl_slist_append(requestData->headers, "Content-Type: application/json");
        curl_easy_setopt(curl_handle, CURLOPT_POSTFIELDS, payloadChars);
    }
    curl_easy_setopt(curl_handle, CURLOPT_HTTPHEADER, requestData->headers);
    curl_easy_setopt(curl_handle, CURLOPT_CUSTOMREQUEST, method.c_str());
    curl_easy_setopt(curl_handle, CURLOPT_PRIVATE, requestData);

    return curl_handle;
}

bool SendsViaGraph(shared_ptr<Account> account) {
    return account->provider() == "office365" &&
        (account->SMTPUsername() == "" || account->SMTPUsername() == account->IMAPUsername());
}

static const string GRAPH_ME = "https://graph.microsoft.com/v1.0/me";
// Graph docs: single requests for files under 3 MB, upload sessions (PUT ranges < 4 MB) above.
static const size_t GRAPH_SINGLE_REQUEST_MAX = 3 * 1024 * 1024;
static const size_t GRAPH_UPLOAD_CHUNK = 12 * 320 * 1024;

// One Graph HTTP call. token == "" sends no Authorization (upload URLs are pre-authenticated).
// Throws SyncException on non-2xx; the handle is cleaned up either way.
static json graphCall(string method, string url, string token, const string & body, string contentType, string extraHeader = "") {
    CURL * curl_handle = curl_easy_init();
    CurlRequestData *requestData = new CurlRequestData();
    curl_easy_setopt(curl_handle, CURLOPT_PRIVATE, requestData);
    if (token != "") {
        requestData->headers = curl_slist_append(requestData->headers, ("Authorization: Bearer " + token).c_str());
    }
    // "Content-Type:" with no value stops curl adding its form-urlencoded default
    requestData->headers = curl_slist_append(requestData->headers, ("Content-Type: " + contentType).c_str());
    requestData->headers = curl_slist_append(requestData->headers, "Expect:");
    if (extraHeader != "") {
        requestData->headers = curl_slist_append(requestData->headers, extraHeader.c_str());
    }
    curl_easy_setopt(curl_handle, CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl_handle, CURLOPT_CONNECTTIMEOUT, 20);
    curl_easy_setopt(curl_handle, CURLOPT_CUSTOMREQUEST, method.c_str());
    curl_easy_setopt(curl_handle, CURLOPT_HTTPHEADER, requestData->headers);
    curl_easy_setopt(curl_handle, CURLOPT_POSTFIELDS, body.data());
    curl_easy_setopt(curl_handle, CURLOPT_POSTFIELDSIZE, (long)body.size());
    return PerformJSONRequest(curl_handle);
}

void GraphSendMIME(shared_ptr<Account> account, mailcore::Data * mime) {
    try {
        auto parts = SharedXOAuth2TokenManager()->partsForAccount(account, "https://graph.microsoft.com/Mail.Send");
        graphCall("POST", GRAPH_ME + "/sendMail", parts.accessToken, mime->base64String()->UTF8Characters(), "text/plain"); // 202
    } catch (SyncException & ex) {
        throw SyncException("send-failed", "Microsoft Graph sendMail: " + ex.key + " " + ex.debuginfo, false);
    }
}

void GraphSendMessage(shared_ptr<Account> account, mailcore::MessageBuilder & builder) {
    mailcore::Data * mime = builder.data();
    if (mime->length() < GRAPH_SINGLE_REQUEST_MAX) {
        GraphSendMIME(account, mime);
        return;
    }

    // Too big for one request: create a draft without the file attachments, attach each one
    // (upload session if >= 3 MB), then send the draft. Inline images stay in the MIME.
    mailcore::Array * files = builder.attachments();
    mailcore::Data * bare = mime;
    if (files != NULL && files->count() > 0) {
        files->retain();
        builder.setAttachments(NULL);
        bare = builder.data();
        builder.setAttachments(files); // stores a copy
        files->release();
        files = builder.attachments();
    }
    if (bare->length() >= GRAPH_SINGLE_REQUEST_MAX) {
        throw SyncException("send-failed", "Microsoft Graph: message body and inline images exceed 3 MB.", false);
    }

    string token;
    string draftUrl;
    try {
        token = SharedXOAuth2TokenManager()->partsForAccount(account, "https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send").accessToken;
        json draft = graphCall("POST", GRAPH_ME + "/messages", token, bare->base64String()->UTF8Characters(), "text/plain");
        if (!draft.count("id") || !draft["id"].is_string()) {
            throw SyncException("send-failed", "Microsoft Graph returned no draft id: " + draft.dump(), false);
        }
        CURL * esc = curl_easy_init();
        char * id = curl_easy_escape(esc, draft["id"].get<string>().c_str(), 0);
        draftUrl = GRAPH_ME + "/messages/" + id;
        curl_free(id);
        curl_easy_cleanup(esc);
        spdlog::get("logger")->info("-- Graph draft created (internetMessageId {}), adding {} attachments", draft["internetMessageId"].is_string() ? draft["internetMessageId"].get<string>() : "", files->count());

        for (unsigned int i = 0; i < files->count(); i++) {
            mailcore::Attachment * a = (mailcore::Attachment *)files->objectAtIndex(i);
            mailcore::Data * d = a->data();
            string name = a->filename() ? a->filename()->UTF8Characters() : "attachment";
            string type = a->mimeType() ? a->mimeType()->UTF8Characters() : "application/octet-stream";
            size_t len = d->length();
            if (len < GRAPH_SINGLE_REQUEST_MAX) {
                json att = {{"@odata.type", "#microsoft.graph.fileAttachment"}, {"name", name}, {"contentType", type}, {"contentBytes", d->base64String()->UTF8Characters()}};
                graphCall("POST", draftUrl + "/attachments", token, att.dump(), "application/json");
                continue;
            }
            json req = {{"AttachmentItem", {{"attachmentType", "file"}, {"name", name}, {"size", len}, {"contentType", type}}}};
            json session = graphCall("POST", draftUrl + "/attachments/createUploadSession", token, req.dump(), "application/json");
            string uploadUrl = session.value("uploadUrl", "");
            for (size_t off = 0; off < len; off += GRAPH_UPLOAD_CHUNK) {
                size_t n = min(GRAPH_UPLOAD_CHUNK, len - off);
                string range = "Content-Range: bytes " + to_string(off) + "-" + to_string(off + n - 1) + "/" + to_string(len);
                graphCall("PUT", uploadUrl, "", string(d->bytes() + off, n), "application/octet-stream", range);
            }
        }
        graphCall("POST", draftUrl + "/send", token, "", ""); // 202
    } catch (SyncException & ex) {
        if (draftUrl != "") {
            try { graphCall("DELETE", draftUrl, token, "", ""); } catch (SyncException &) {} // don't leave a half-built draft
        }
        throw SyncException("send-failed", "Microsoft Graph large send (needs Mail.ReadWrite consent): " + ex.key + " " + ex.debuginfo, false);
    }
}

const string PerformRequest(CURL * curl_handle) {
    string result;
    curl_easy_setopt(curl_handle, CURLOPT_WRITEFUNCTION, _onAppendToString);
    curl_easy_setopt(curl_handle, CURLOPT_WRITEDATA, (void *)&result);
    CURLcode res = curl_easy_perform(curl_handle);
    ValidateRequestResp(res, curl_handle, result);
    return result;
}


const string PerformExpectedRedirect(string url) {
    CURL * curl_handle = curl_easy_init();
    curl_easy_setopt(curl_handle, CURLOPT_URL, url.c_str());
    curl_easy_setopt(curl_handle, CURLOPT_CONNECTTIMEOUT, 10);

    string result;
    curl_easy_setopt(curl_handle, CURLOPT_WRITEFUNCTION, _onAppendToString);
    curl_easy_setopt(curl_handle, CURLOPT_WRITEDATA, (void *)&result);
    curl_easy_perform(curl_handle);

    long http_code = 0;
    string redirect = "";
    curl_easy_getinfo(curl_handle, CURLINFO_RESPONSE_CODE, &http_code);
    // RFC 6764 §5: clients MUST follow redirects on .well-known. 308 postdates the RFC.
    bool isRedirect = http_code == 301 || http_code == 302 || http_code == 303 || http_code == 307 || http_code == 308;
    if (isRedirect) {
        char * _redirect = nullptr;
        if (curl_easy_getinfo(curl_handle, CURLINFO_REDIRECT_URL, &_redirect) == CURLE_OK && _redirect != nullptr) {
            redirect = string(_redirect);
        }
    }
    curl_easy_cleanup(curl_handle);
    return redirect;
}


const json PerformJSONRequest(CURL * curl_handle) {
    string result = PerformRequest(curl_handle);
    json resultJSON = nullptr;
    try {
        resultJSON = json::parse(result);
    } catch (json::exception &) {
        resultJSON = {{"text", result}};
    }
    CleanupCurlRequest(curl_handle);
    return resultJSON;
}

void ValidateRequestResp(CURLcode res, CURL * curl_handle, string resp) {
    char * _url = nullptr;
    if (curl_easy_getinfo(curl_handle, CURLINFO_EFFECTIVE_URL, &_url) != CURLE_OK || _url == nullptr) {
        throw SyncException(res, "Unable to get URL");
    }
    string url { _url };

    if (res != CURLE_OK) {
        CleanupCurlRequest(curl_handle);
        throw SyncException(res, url);
    }

    long http_code = 0;
    curl_easy_getinfo(curl_handle, CURLINFO_RESPONSE_CODE, &http_code);

    if (http_code < 200 || http_code > 209) {
        CleanupCurlRequest(curl_handle); // note: cleans up _url;
        
        bool retryable = ((http_code != 403) && (http_code != 401));
        if (resp.find("invalid_grant") != string::npos) {
            retryable = false;
        }
        
        string debuginfo = url + " RETURNED " + resp;
        throw SyncException("Invalid Response Code: " + to_string(http_code), debuginfo, retryable);
    }
}

// Shorthand methods for Identity Server

CURL * CreateIdentityRequest(string path, string method, const char * payloadChars) {
    string url { MailUtils::getEnvUTF8("IDENTITY_SERVER") + path };

    if (Identity::GetGlobal() == nullptr) {
        // Note: almost all the backend APIs require auth except for /api/resolve-dav-hosts
        // and calls to this method should be avoided if no identity is present.
        return CreateJSONRequest(url, method, "", payloadChars);
    }

    string plain = Identity::GetGlobal()->token() + ":";
    string encoded = MailUtils::toBase64(plain.c_str(), strlen(plain.c_str()));
    string authorization = "Basic " + encoded;
    return CreateJSONRequest(url, method, authorization, payloadChars);
}

const json PerformIdentityRequest(string path, string method, const json & payload) {
    string payloadString = payload.dump();
    return PerformJSONRequest(CreateIdentityRequest(path, method, payloadString.c_str()));
}
