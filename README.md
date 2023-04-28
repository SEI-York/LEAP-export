# LEAP-export
This script is intended to extract data from LEAP model into UNFCCC reporting format.
# How to?
- Add the category code in the template file to LEAP branch as a tag, for example, UNFCCC_1.A.1.
    - NOTE: please make sure to untick the "Tag children" option in Tag's Settings
- Download the [script file](https://raw.githubusercontent.com/SEI-York/LEAP-export/main/Export%20to%20UNFCCC%20template.vbs) and the [template file](https://github.com/SEI-York/LEAP-export/raw/main/UNFCCC%20template.xlsx)
- Place the [template file](https://github.com/SEI-York/LEAP-export/raw/main/UNFCCC%20template.xlsx) in this folder `LEAP_AREA/_Settings/_DictionaryNX`
- Run the [script file](https://raw.githubusercontent.com/SEI-York/LEAP-export/main/Export%20to%20UNFCCC%20template.vbs) with LEAP's Script Editor 
    - `Advanced -> Edit scripts` and choose the downloaded script file
    - Please update the fuel list according to the settings in LEAP