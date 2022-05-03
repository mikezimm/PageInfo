# fps-page-info

## Summary
Web part to combine both Header based Table of Contents and Page Info (meta data) into same web part


## Setup and install
yo @microsoft/sharepoint
npm install
npm install react-json-view

Copied in AdvancedPageProperties
did general npm install here hoping to get @pnp but it did not work
npm install @pnp/sp
npm uninstall @pnp/sp //Had errors

Then did this
npm install @pnp/logging @pnp/common @pnp/odata @pnp/sp --save

```cmd
[21:24:35] Starting subtask 'tsc'...
[21:24:35] [tsc] typescript version: 3.9.10
[21:24:36] Finished subtask 'copy-static-assets' after 2.12 s
[21:24:38] Finished subtask 'tslint' after 2.73 s
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(10,115): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(10,155): error TS1005: ';' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(10,156): error TS1131: Property or signature expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(10,157): error TS1109: Expression expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(11,5): error TS1128: Declaration or statement expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(11,101): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(11,112): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(11,127): error TS1109: Expression expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(12,5): error TS1128: Declaration or statement expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(12,115): error TS1005: '(' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(13,5): error TS1128: Declaration or statement expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(13,117): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(13,132): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(13,150): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(13,157): error TS1109: Expression expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(14,5): error TS1128: Declaration or statement expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(14,95): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(14,108): error TS1005: ',' expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(14,115): error TS1109: Expression expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(15,5): error TS1128: Declaration or statement expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(15,85): error TS1109: Expression expected.
[21:24:38] Error - [tsc] node_modules/@pnp/queryable/queryable.d.ts(16,1): error TS1128: Declaration or statement expected.
[21:24:38] Error - 'tsc' sub task errored after 2.81 s
 exited with code 2
 ```