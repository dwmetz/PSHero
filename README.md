# PSHero

PSHero Powershell Menu v1.0
https://github.com/dwmetz/PSHero/
All rights to 3rd party scripts remain with the original owners

Note: do a find/replace for D:\PowerShell\PSHero\ and subsititute the path where PSHero scripts are locally stored

## Background

We have a collection of internal scripts that we use frequently. As more scripts (or scriptlets) are added to the frequently used, I wanted a means to expose all the scripts to the team and to put some organization to it. I also wanted to easily support changes or additions to the referenced scripts. What I wound up building was a simple PowerShell menu structure.

## Contents
The scripts included in this demo menu include:

- Launch PowerShell with an alternate credential
- Login to O365, Legacy and Modern Auth.
- Bitlocker (AD) retrieval
- Host profiling and an attention getting PING
- IRMemPull (Memory Acquisition) – https://github.com/n3l5/irMempull
- A script that adds an examiner permissions to subject mailbox (for Magnet Axiom collection)
- MX header parsing
- Sadphishes – https://github.com/EdwardsCP/powershell-scripts/blob/master/SADPhishes.ps1
- Convert Unix time to human readable
