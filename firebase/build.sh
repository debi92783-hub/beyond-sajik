#!/bin/sh
# 올릴 파일 묶음 만들기
#  - public/ : 업무 화면 (admin.html → public/index.html) → beyond-sajik.web.app
#  - home/   : 메인화면 (index.html → home/index.html)   → beyond-sajik-home.web.app
cd "$(dirname "$0")" && rm -rf public home && mkdir -p public/img home/img \
 && cp ../admin.html public/index.html && cp ../logo.png public/ && cp ../img/entrance.jpg public/img/ \
 && cp ../index.html home/index.html && cp -r ../img/curriculum ../img/method ../img/dream-scene.svg home/img/
