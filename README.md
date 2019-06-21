# Uninett SPFx SharePoint Web Parts

A collection of web parts developed by Uninett.

## User Directory

Displays users from your Office 365 tenant in a customised [DetailsList](https://developer.microsoft.com/en-us/fabric#/controls/web/detailslist) component.  

The web part uses MS Graph REST API to retrieve Office 365 users from Azure AD.  
Using either the built-in search box or the Search Box web part, you can then filter users by name or department.

### Features

- Sort columns by clicking on their header
- Select which columns to display
- Create custom column headers
- Resize columns
- Customise the visual presentation of the results

### Properties

- **API:** The MS Graph query for the group of users you want to display (e.g. *users* for all users, or *groups/{group_ID}/members* for users in a specific group).
- **Compact mode:** Toggle between compact and normal row spacing.
- **Row colour:** Toggle between single row colour (white) and alternating row colours (blue and white).
- **Search box source:** Toggle between using the built-in search box and the Search Box web part. If you want to use the Search Box web part, simply add it to the page and it will automatically connect to the User Directory.
  - **Placeholder text:** Specify the placeholder text that is shown when the search box is empty (for built-in search box).
- **Select columns:** Use the checkboxes to select the columns you want to display in the User Directory.
- **Custom column headers:** Lets you create custom column headers. Leave the fields blank to use the default column headers.

<p align="center">
  <img src="/readme-images/user-directory-demo.JPG" alt="User Directory demo"/>
</p>

## Search Box

Simply add this to a page with a User Directory on it to be able to search through the directory.

This web part uses the [ReactiveX (RxJs)](http://reactivex.io/) library to allow communication between the two web parts.  
Its functions are identical to the built-in search box of User Directory.

### Properties

- **Placeholder text:** Specify the placeholder text that is shown when the search box is empty.

<p align="center">
  <img src="/readme-images/search-box-demo.JPG" alt="Search Box demo"/>
</p>

## Social Links

Web part that contains icons, with links, to social media. Property to change background color and toggle icon color between black and white. If URL to a social media is empty then the icon is removed. Supports:
- Facebook
- Twitter
- LinkedIn
- YouTube

<p align="center">
  <img src="/readme-images/social-links-demo.JPG" alt="Social links demo"/>
</p>

## FlipClock

Flip clock that lets you toggle between two modes:

- **Time**
  - Simply displays the current time.
  - Properties let you select time format (24h/12h) and wether to show seconds or not.
- **Weekly countdown**
  - Counts down days, hours, minutes and seconds until a specified day and time in the upcoming week.
  - Properties let you select the day of week and time of day, as well as optional text to include under the clock.

General properties let you switch between the modes, change clock size and toggle between black and white text color.

<p align="center">
  <img src="/readme-images/flip-clock-demo.JPG" alt="FlipClock demo"/>
</p>

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
