Insanity Uploader
==============

This Ruby script adds projects to the database of projects maintained in the
[802.1 Maintenance Database](https://github.com/jlm/maint), and updates them.
Data can be drawn from the Insanity Spreadsheet which the script can parse,
from the IEEE Standards myProject system which the script can log in to and query, and from the
802.1 Email reflector which the script can parse for ballot announcements. 
It then parses and uploads these to the web application.  The maintenance
database is a Ruby on Rails web app which exposes a JSON API as well as a web-based user interface.

The email archive is managed using Listserv which generates an index on HTML pages.
The excellent XML and HTML parser, Nokogiri, can parse these without difficulty.

Configuration
-------------
The URLs of the maintenance databse API and the Listserv message archive, together with the usernames
and passwords for each, are stored in `secrets.yml` which is not included in the sources.
An example of that file file is available as `example-secrets.yml`.

Command-Line Options
--------------------
Without any options, the script just logs in to the maintenance database and exits.
* `--filepath insanity.xlsx`  (abbreviated `-f`): Parse the Insanity spreadsheet in the file.  This looks at three tabs
  in the spreadsheet, `People`, `TaskGroups` and `Projects`.  The `People` sheet is optional - see the `--people` option
  for more information.  The `TaskGroups` tab must contain a list of the abbreviations used in the
  `TG` column of the `Projects` tab.  The first column is the TG abbreviation. The second column is
  the Task Group name.  The third and fourth columns, if present, are the first and last names of the
  Task Group chair, and are used by the `--task-groups` option.
  With the `--filepath` option, the script will parse each line of the `Projects` tab and create a new project in the
  database for each project.  Dated items are converted into events and entered into the list of events for the
  new project.  The `Editor` column is currently not proecessed.
* `--people` (abbreviated `-p`): when given with the `--filepath` option, this causes the list of people in the `People`
  tab to be converted to entries in the database.  The columns in order are Role, First name, Last name, Email and
  Affiliation.  A unique record consists of a combination of Role, First name and Last name.  That is, a single
  individual can have multiple entries, one for each role they fulfil, such as Editor, Chair or ViceChair.  People must
  exist in the database before they can be entered as Chair of a Task Group for example.
* `--task-groups` (abbreviated `-t`): when given with the `--filepath` option, create (or update) Task Groups in the
  database based on the information on the TaskGroups tab of the Insanity spreadsheet.  Task Groups must exist 
  in the database before projects can be assigned to them.
* `--delete-existing` (abbreviated `-d`): delete an existing project before creating its replacement.
* `--update` (abbreviated `-u`): update existing projects with information from the Insanity spreadsheet.
* `--devserv` (abbreviated `-b`): log in to the IEEE myProject system and scan the list of active PARs with designations
  starting '802[a-zA-Z]' and '802.1[a-zA-Z]'.  Each project is added to or updated in the database.  The original PAR URL
  and full title are recorded.  The current PAR is examined and the PAR dates are entered as project events. 
* `--par-report` (abbreviated `-r`): Parse the PAR Report from MyProject. The --par-report option takes a filename
  argument. The file is a set of lines representing the projects to look for in the PAR Report.  Each line has the
  project designation, a colon, a space, and then the abbreviated Task Group name to which to assign the project.
  The task group has to exist. The entire PAR Report is parsed.  When a project is found which matches an entry in the
  list described above, The project is added (or updated) in the database using information from the PAR Report entry
  and from the linked HTML-format PAR.
* `--mailserv` (abbreviated `-m`): Log in to the 802.1 email archive and scan it for messages announcing task group and
  working group ballots.  The ballot dates are recorded in the project event list.
* `--secrets secretfile.yml` (abbreviated `-s`): Use the given file for the configuration options rather than the
  default `secrets.yml`. 

Usage
-----
Initially it is envisaged that people, task groups and projects are imported from the Insanity spreadsheet, but this is
not required.  The information can be entered into the Web interface directly, instead.  Subsequently the idea is that 
the script is run regularly to update project information from the myProject server and the mail server.

Deployment
----------

The script can be deployed in a Docker container.  I use a very simple one based on Ruby:2.3.0-onbuild.
This method is frowned upon. Bear in mind that the `secrets.yml` file will be included into the
container, so the container is secret too.  There are methods to isolate the secret information
from the container, but I have not bothered to do this.

License
-------
Copyright 2016-2017 John Messenger

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Author
------
John Messenger, ADVA Optical Networking Ltd., Vice-chair, 802.1
