#!/bin/sh

platform=linux-x64
dir=linux

cd $(dirname "$0")
cd ../

dotnet publish -c Release --self-contained true -r $platform -o output/${dir}
