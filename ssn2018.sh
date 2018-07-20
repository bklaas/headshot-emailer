#!/bin/bash

# wrapper script for SSN2018 conference to invoke headshot emailer

subject="SSN Discover 2018 National Conference: Your Professional Photo is attached"
imagedir="foo"
cmd="python headshot_emailer.py \"${subject}\" \"${imagedir}\" \"email_content/ssn2018.html\" -d"
echo $cmd

