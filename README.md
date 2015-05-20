Daily Scrum add-on
==================

This repo is a Google Spreadsheet add-on that helps to send daily mails.

Install
-------

1. In a Google Spreadsheet, in the menu click on: "add-on" ("Modules complémentaires") > "Get add-ons" ("Télécharger des modules complémentaires")
2. In the top left select list click on "For theodo.fr"
3. Accept the permissions (this add-on requires many permissions... I guaranty the code comes from this repo but you can also ask for view on [the add-on source-code](https://script.google.com/macros/d/MdMf53huZLjmmqvprh6GzO5aOUM6bORhY/edit?uiv=2&mid=ACjPJvEjPflgp998EriuzN0cdXDnU01i1f1i26YWd-Y_rtAfFlbk2nJ0KULlYpbI6TvAJ1b9tMx9f6TgGmtu9h10uszMEKU9z8S6pSdQIUBg16gC9EY1xENPZVAt6ZveV3TAUKiGGoXqQcc)).

Usage
-----

1. Start a new project (menu > add-on > daily scrum > new project)
2. You might fill the tech' team (it will pre-fill the next form)
3. Create a new sprint sheet (menu > add-on > daily scrum > new sprint)
4. Fill the greenish cells
5. Send daily mails:
  - Write a template of your daily mail as a draft in your Gmail (an example of template with all dynamics fields is available in menu > add-on > daily scrum > daily mail). Don't forget to fill the recipient and the cc field.
  - Click on menu > add-on > daily scrum > daily mail and pick the corresponding draft in the list of drafts
  - Fill the number of points your PO has to validate
  - Preview the daily mail and send it!

Daily template
--------------

In your Gmail draft template, you can use the following variables:
 - {sprintNumber}
 - {sprintGoal}
 - {sprintDay}
 - {date} (current date)
 - {totalPoints}
 - {donePoints}
 - {toValidatePoints}
 - {toStandardPoints}
 - {earlyOrLate} ('Avance' ou 'Retard')
 - {doneColorS} {doneColorE} (html span tag to have dynamic colors)
 - {validationColorS} {validationColorE} (html span tag to have dynamic colors)
 - IMG https://scrutinizer-ci.com/g/theodo/backtest-api/badges/build.png IMG (to get a screenshot of the image at the correponding url)

To send a daily mail you should edit your draft every morning without caring about repetitive work (copy the previous mail, update the subject, get the BDC, update the points).

The most useful thing to me is the IMG tag that fetch a url and display the corresponding image. I use it to display a screenshot of the continuous integration badges and code quality badges.
