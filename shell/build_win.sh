#!/bin/sh

platform=win-x64
dir=win

cd $(dirname "$0")
cd ../

dotnet publish -c Release --self-contained true -r $platform -o shell/output/${dir}
