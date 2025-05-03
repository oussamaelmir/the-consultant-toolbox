#!/usr/bin/env bash
npm ci
npm run build
rm -rf "$HOME/site/wwwroot"/*
cp -R dist/* "$HOME/site/wwwroot/"
