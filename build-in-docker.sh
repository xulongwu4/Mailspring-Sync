#!/bin/bash
set -e
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq autoconf automake build-essential clang cmake git libc-ares-dev libctemplate-dev libcurl4-openssl-dev libglib2.0-dev libicu-dev liblzma-dev libsasl2-dev libssl-dev libsasl2-modules libtidy-dev libtool libxml2-dev uuid-dev pkg-config >/dev/null
cp -a /src /build
cd /build/Vendor/libetpan && ./autogen.sh >/dev/null && make -j$(nproc) >/dev/null && make install prefix=/usr >/dev/null
cd /build/Vendor/mailcore2 && rm -rf build && mkdir build && cd build && cmake .. >/dev/null && make -j$(nproc) MailCore >/dev/null
cd /build && rm -f CMakeCache.txt && cmake . >/dev/null && make -j$(nproc) 2>&1 | grep -E "error|warning: unused|Error" || true
cp /build/mailsync /out/mailsync.bin
echo BUILD_DONE
