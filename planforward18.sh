#!/bin/bash
# wrapper script for PlanForward18 conference to invoke headshot emailer
conf='planforward18'
./headshot_emailer "$@" --conference ${conf}

