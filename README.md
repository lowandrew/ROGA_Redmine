# Redmine ROGA Automation

Running these scripts will allow you to have ROGA reports generated in an automated fashion, with requests generated
through Redmine.

## Downloading and Installing

- Clone this repository: `git clone --recursive https://github.com/lowandrew/ROGA_Redmine.git`
- Move into the repository: `cd ROGA_Redmine`
- Install python requirements: `pip3 install -r requirements.txt`


## Running the redmine automation script
- From the ROGA_Redmine directory `python ROGA_Redmine_Run.py`
- Accept default settings where possible, put in API Key (found in 'My Account' on Redmine.)

## Running requests through redmine
- Issue subject must be AutoROGA 
- Attach an up-to-date copy of ROGA_summary_OLF.xlsx or ROGA_summary_OLC.xlsx
- First line of the description has to be either Listeria, VTEC, Salmonella, or Salmonella_Feed, depending on which
which type of report you want to run
- Subsequent lines of the description should be the seqIDs needed for the ROGA

## Getting Supervisor Set Up
I'll have to figure out how to do this at some point soon...



