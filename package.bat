set PACKAGE_NAME=KiTTY

rmdir /q /s build
mkdir build

cd src
zip ..\build\%PACKAGE_NAME%.keypirinha-package *.py *.ini *.ico lib/*.py icons/*.*

cd ..