# Uninett SPFx SharePoint Web Parts

Contains currently only one web part.

## Social Links

Web part that contains icons, with link, to social media. Property to change background color and toggle icon color between black and white. If URL to a social media is empty then the icon is removed. Supports:
- Facebook
- Twitter
- LinkedIn
- YouTube

![spfx-uninett-webparts-social_links](/readme-images/spfx-uninett-webparts-social_links.jpg)

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
