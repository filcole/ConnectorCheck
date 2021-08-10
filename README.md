# Deprecated Actions

## Program Command line

Unpacked solution folder

## Program Steps

0. Download list of deprecated actions from tinternets

1. Search workflows folder to find Power Automate files \*.json

   1. Capture connection references and connectors used
   2. Capture actions
   3. See if actions are deprecated and list

2. Find MSAPP
   1. Find Canvas Apps
   2. Explode using MSAPP
   3. Capture connection references and connectors used
   4. Capture PowerFx formulas using these connection references
   5. See if any references are made to deprecated actions

## Desired Output

1. list of flows with deprecated actions one row per flow actions with

   - flowfilename
   - flowname
   - connectionref
   - connectoruniquename
   - connectorname
   - action
   - actionname
   - actiondescription

2. Desired Output
   - Name of canvas app
   - Connector
   - action
