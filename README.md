# Getting started with OfficeJs add-in

## Introduction

A starting point for Angular based Word Office JS add-in with a shared runtime enabled.

Availble functionality:
- Taskpane with a single button that adds a paragraph with text
- Ribbon button that adds a paragraph with text
- Ribbon button wich hilights selected text, and gets disabled when no text is selected (to demonstrate that the shared runtime is working)

## Set up environment

Install nvm:

    winget install CoreyButler.NVMforWindows

Install lts version of node using commands:

    nvm install lts
    nvm use lts

Install Angular CLI

    npm install -g @angular/cli

## Set up VS Code for Debugging

Following VS Code add-ins are already set as workspace recommendations and should install automatically. Please use them!

- [Microsoft Edge Tools for VS Code](https://marketplace.visualstudio.com/items?itemName=ms-edgedevtools.vscode-edge-devtools)
- [Microsoft Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger)

## Build

    npm install

## Certs

    mkdir .cert
    npm run create-certs
    Import-Certificate -FilePath .cert/ca.crt -CertStoreLocation 'Cert:\CurrentUser\Root'

## Debug

F5
