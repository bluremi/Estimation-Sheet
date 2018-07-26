# Estimation-Sheet
This is a test repo to set up a version control system for the VBA modules in an Excel spreadsheet.

The spreadsheet includes vba code in both worksheets and in separate modules.

The workflow depends on githooks running a python script to extract the VBA code into separate files, which can then be imported back into Excel when ready to be published.
## Prerequisites
Although the instructions below include commands for Windows, they have only been tested on Mac.
### Python 2.7 or higher
Install Python 2.7 or higher
### Python oletools package
Linux, Mac OSX, Unix
To download and install/update the latest release version of oletools, run the following command in a shell:

```
sudo -H pip install -U oletools
```
Important: Since version 0.50, pip will automatically create convenient command-line scripts in /usr/local/bin to run all the oletools from any directory.

Windows
To download and install/update the latest release version of oletools, run the following command in a cmd window:
```
pip install -U oletools
```
Important: Since version 0.50, pip will automatically create convenient command-line scripts to run all the oletools from any directory: olevba, mraptor, oleid, rtfobj, etc.

## Setup
Clone the repository
```
git clone https://github.com/bluremi/Estimation-Sheet.git
```

Copy the files `pre-commit.py` and `pre-commit` into your `.git/hooks` directory within the cloned repo.

Run the following command to make sure the hooks are triggeed whenever you do a `commit`:
```
chmod +x .git/hooks/pre-commit
```

From now on, when you execute `git commit`, the hook extracts the VBA code and adds it to your commit as a series of .bas files.

