#!/bin/bash

# wrapper script for PlanForward18 conference to invoke headshot emailer

subject="Investacorp PlanForward18 National Conference: Your Professional Photo is attached"
imagedir="foo"
cmd="python headshot_emailer.py \"${subject}\" \"${imagedir}\" \"email_content/planforward18.html\" -d"
echo $cmd

