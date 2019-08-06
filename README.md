# Uninett SPFx SharePoint Web Parts

A collection of web parts developed by Uninett.

* [Site Directory](documentation/readme/site-directory.md)
* [Create Site Button](documentation/readme/create-site.md)
* [User Directory](documentation/readme/user-directory.md)
* [Search Box](documentation/readme/search-box.md)
* [Social Links](documentation/readme/social-links.md)
* [Flip Clock](documentation/readme/flip-clock.md)


___

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
