##################################################################################################
#
# 			-- Sample GlobalSiege server configuration file --
#
# When passed to GlobalSiege by command line, or by drag and drop to start the app, Globalsiege
# will read this file and parse the commands listed below during startup. This file was designed
# to allow GlobalSiege to run as a headless online war server. To run GlobalSiege as a scheduled
# task, enter the following command in the Task - Run field: 
# C:\PROGRA~1\GLOBAL~1\GlobalSiege.exe "C:\Program Files\GlobalSiege\Sample GlobalSiege.conf"
# You will, of course, modify and rename the file to something sensible.
#
# The hash '#' character at the start of a line indicates a comment and is ignored. Blank lines
# are also ignored. Command syntax is line delimited with a general format of "Command_Name=Value".
# The Value cannot span multiple lines. Use "$0D$0A" for line breaks, "$3D" for equals the sign "="
# and "$24" for a dollar "$" sign.
#
##################################################################################################
#
# If Server_Mode is set to TRUE, GlobalSiege will run in a CPU efficient way, this means minimal 
# graphics, all sound off. GlobalSiege will attempt to begin a new Internet session. 
# Options: [TRUE|FALSE]

Server_Mode=TRUE

##################################################################################################
#
# If Headless is set to TRUE, the display is not refreshed at all. Use this setting for running
# session hosts in the background. Uncomment to activate by removing the hash '#'.
# Options: [TRUE|FALSE]

#Headless=TRUE

##################################################################################################
#
# The number of seconds to delay the GlobalSiege start up. This is needed if there is more than one 
# host session started from the task  scheduler. Sessions that are posted from the same IP address at  
# the same time confuses the Internet Indexing server.
# Options: [0-9999]

#start_delay=60

##################################################################################################
#
# Server_ID is a unique number given to a GlobalSiege instance to assist differentiate log files
# when multiple instances are running.
# Values: [00-99]

Server_ID=01

##################################################################################################
#
# Log_Level controls the logging detail level. The higher the number, the more detail is logged.
# Values can be from 0 to 9.
# Values: [0-9]

Log_Level=0

##################################################################################################
#
# The name of the war file to open. This can be just a war file name in which case GlobalSiege will
# search its normal war file locations or it could be a full path to war file in any location.

War_File=Default Online Host

##################################################################################################
#
# Automatically restart the war after a short delay.

Auto_Restart=TRUE

##################################################################################################
#
# The name of the session. Use "3D" in place of the equal sign "=".

Session_Name=Default Online War

##################################################################################################
#
# The main port. These port numbers must be different for each session host running on the same server.

TCP_Port=12345

##################################################################################################
#
# The broadcast port. These port numbers must be different for each session host running on the same server.

UDP_Port=12345

##################################################################################################
#
# The maximum number of armies that each remote terminal can claim.

Max_Armies_Per_Terminal=2

##################################################################################################
#
# The maximum number of connections from each IP address.

Max_Connections_Per_IP=2

##################################################################################################
#
# Set the session password. Leave blank if you do not want a password set.

Session_Password=

##################################################################################################
#
# The time limit in seconds that of no activity from each player. If a player does not make any
# move within this time limit, their turn will be forfieted. Set to 0 for no time limit.
Turn_Time_Limit=60

##################################################################################################
#
# The welcome message to display in the session locator and to the player upon initial connection.
# Use the following character codes for the following non alphnumeric characters:
# "=" 	-> $3D
# <CR>	-> $0D$0A
# "$"	-> $24
# tab	-> $09
# Any other character can be included by appending a dollar sign "$" in front of the character's 
# two digit hexadecimal ASCII code. For example, "Session = $money" would be written as
# "Session $3D $24money"

Welcome_Message=Welcome to my online session.

##################################################################################################
#
# The speed that the computer players take their turns.
# [TRUE|FALSE]

Fast_War=TRUE

##################################################################################################

##################################################################################################
#
# Turn Clem's counter on if TRUE.
# [TRUE|FALSE]

Counter_On=TRUE

##################################################################################################
