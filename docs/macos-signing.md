# macOS Signing Notes

Quick Document can build an unsigned DMG for local testing with:

```bash
npm run dist:mac
```

For public distribution without Gatekeeper warnings, the Mac build needs an active Apple Developer Program certificate:

- Required: `Developer ID Application`
- Optional for pkg workflows: `Developer ID Installer`

The iOS certificates commonly named `iPhone Developer`, `iPhone Distribution`, `Apple Development`, or `Apple Push Services` cannot sign a macOS Electron app for outside-the-App-Store distribution.

Once a valid `Developer ID Application` identity is installed in the login keychain, `electron-builder` can auto-detect and sign the app during `npm run dist:mac`.

For notarization, add Apple notarization credentials in CI or the local environment after signing is available.
