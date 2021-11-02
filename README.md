# bulk-pdf-generator-with-template-and-mailer

Just go to your spreadsheet where your dada is stored >> then click on tools >> then click on Script Editor >> a new page will open there paste this code and make changes acording to your needs

to change data dinamically in the slide template write the fields link this "{{name}}"

and then in the script find the string from the slide and replace it with your preferable string like this one

``body.replaceAllText("{{name}}",values[i][0]);``

it is located in line no 11 in script

Then run it and it will do your work
