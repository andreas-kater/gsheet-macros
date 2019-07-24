#!/bin/sh

set -e
set -o pipefail

#create project folder
if [ -d "$1" ]; then
  cd "$1"
  if [ -L "$1" ]; then
    echo "there is already a symlink named $1"
  else
    echo "there is already a directory named $1"
  fi
else
  mkdir $1
  cd $1
fi

#create clasp project

#set script ID
# echo "ScriptID: "
# read scriptId
# echo "{\"scriptId\":\"$scriptId\"}" > .clasp.json
clasp create --title "$1" --type sheets

#add my standard macros
ln -sf ../../my-macros/macros.js
ln -sf ../../my-macros/appsscript.json
ln -sf ../../my-macros/stitch-sync.js

#push to project
clasp push --force

#make sure you're in the newly created directory
exec zsh

