#!/bin/sh
# 올릴 파일 묶음 만들기: admin.html → public/index.html (+ 로고·사진)
cd "$(dirname "$0")" && rm -rf public && mkdir -p public/img && cp ../admin.html public/index.html && cp ../logo.png public/ && cp ../img/entrance.jpg public/img/
