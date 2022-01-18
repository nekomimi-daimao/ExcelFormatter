#!/bin/sh

platform=osx-x64
dir=mac

cd $(dirname "$0")
cd ../

dotnet publish -c Release --self-contained true -r $platform -o shell/output/${dir}
