# Uninett SPFx SharePoint Web Parts

Currently contains two web parts.

## SocialLinks

Web part that contains icons, with links, to social media. Property to change background color and toggle icon color between black and white. If URL to a social media is empty then the icon is removed. Supports:
- Facebook
- Twitter
- LinkedIn
- YouTube

![spfx-uninett-webparts-social_links](/readme-images/spfx-uninett-webparts-social_links.jpg)

## FlipClock

Flip clock that lets you toggle between two modes:

- **Time**
  - Simply displays the current time.
 - Properties let you select time format (24h/12h) and wether to show seconds or not.
- **Weekly countdown**
 - Counts down days, hours, minutes and seconds until a specified day and time in the upcoming week.
 - Properties let you select the day of week and time of day, as well as optional text to include under the clock.

General properties let you switch between the modes, change clock size and toggle between black and white text color.

![spfx-uninett-webparts-flipclock](/readme-images/spfx-uninett-webparts-flipclock.jpg)


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
