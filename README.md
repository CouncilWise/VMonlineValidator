# VMonlineValidator
This script validates the XML files that are required to be sent to VMonline.
Simply run the script and open a VMonline formatted XML file.

The script will detect what type of file it is and then process it acordingly folloing VMonline's API data requirements. It can also be ran manually that allows you to pass a '-MissingFields $true/$false' variable.
By default this is set to false meaning that it will only check the data itself to see if it comply's to VMonline's requirements. However if set to $true than it will report any missing fields as an error in the outputed file.

Then the script runs it will export out any errors found to the same directory that the orginiating XML came from.
