#! python3
"""
010323 Purpose of this script is to add signature lines to timesheets for Amanda
It also removes employees pages that she doesn't need to print
It also covers over the footer with a white rectangle, effectively removing it
This will need to be an executable file, she should not have to install python
"""

import os, sys, re
import json
from pypdf import PdfWriter, PdfReader
from datetime import datetime as dt
import traceback
import fitz

# TODO traceback error handling that calls self._debug(). Probably need to remove the try except from __main__

class TS_Format:
    def __init__(self, param_file=str):
        """
        Remove pages containing names from .txt file
        Merge a pdf with another pdf as a watermark
        Add a white rectangle for redacted items
        """

        try:
            # Set variables
            self.scriptpath = os.path.dirname(os.path.realpath(sys.argv[0])) + '\\'
            
            tstamp = dt.now().strftime("%Y-%m-%d_%H-%M-%S")
            vars = self.load_json(param_file)
            self.homepath = 'C:' + os.environ["HOMEPATH"]

            # Replace 'scriptpath' str in json with actual script path
            for key in vars.values():
                try:
                    if key["path"] == 'scriptpath':
                        key["path"] = self.scriptpath
                except TypeError:
                    continue
                except KeyError:
                    continue

            # Find the file to be edited
            searchfile = vars["searchfile"]["namekeyword"]
            searchtype = vars["searchfile"]["type"]
            searchpath = vars["searchfile"]["path"]
            if not searchpath.startswith('C:'):
                searchpath = self.homepath + '\\'+ searchpath +'\\'
            self.editfile = self.find_latest_file(searchfile, searchtype, searchpath)

            # Set the wartermark file
            watermarkpath = vars["mergefile"]["path"]
            if not watermarkpath.startswith('C:'):
                watermarkpath = self.homepath + '\\'+ watermarkpath +'\\'
            self.watermark = watermarkpath + '\\' + vars["mergefile"]["name"]
            
            # Set the output file
            editfilebasename = self.editfile.split("\\")[-1]
            outfilehead = vars["outfile"]['name'] + str(tstamp) + '-'
            outpath = vars["outfile"]["path"]
            if not outpath.startswith('C:'):
                outpath = self.homepath + '\\'+outpath+'\\'
            # Remove spaces and parantheses from the final document name, resolves error
            replace_chars = [' ', '(', ')']
            self.outfile = outpath + outfilehead + editfilebasename
            for c in replace_chars:
                self.outfile = self.outfile.replace(c, "")

            # Get list of names to exclude from final output
            exnamepath = vars["excludednames"]["path"]
            if not exnamepath.startswith('C:'):
                exnamepath = self.scriptpath + exnamepath
            try:
                with open(exnamepath + vars["excludednames"]["name"]) as f:
                    self.exnames = f.read().splitlines()
            except FileNotFoundError:
                sys.exit(input('\nCould not find file "' + vars['excludednames']['name'] + '" in directory ' + exnamepath + '. Ensure file exists in directory.'))

            
            # Set temp file that will be deleted
            self.temp = self.scriptpath + 'temp.pdf'

            # Create a list of pages that contain excluded names (self.exnames)
            self.expages = self.find_excluded_pages()

            # Place white rectangle over footer and create a temp file
            # This has to be done before the watermark - the watermark will replace the rectangle
            self.create_rect()
            
            # Merge the watermark file with the temp file
            # Skip unwanted pages from the list of self.expages and create the final PDF
            self.watermark_merge()

            # Delete temp tile
            os.remove(self.temp)
            
            # Open the new pdf
            os.system(self.outfile)
            
        except Exception as e:
            print('\n\nAn error has occured - Get Kyle!\n\n' + traceback.format_exc())
            vars['debug']['active'] = True

        # Optional debugging. Called if varibles.json ['debug']['active'] = true
        if vars['debug']['active']:
            localvars = None
            if vars['debug']['locals']:
                localvars = locals()
            scope = None
            if vars['debug']['scope']:
                scope = dir()
            self._debug(localvars=localvars, scope=scope, globalvars=vars['debug']['globals'], **vars['debug']['kwargs'])
        
        sys.exit()
        

    def load_json(self, filename=str):
        """Return data from JSON file as a dict"""
        try:
            with open(self.scriptpath + filename) as f:
                data = json.load(f)
        except Exception as e:
            sys.exit('Unable to load JSON data from "' + filename + '". ' + e)
        return data

    def find_latest_file(self, searchfile=str, searchtype=str, searchpath=str):
        """Return the most recent file matching search criteria"""
        allfiles = os.listdir(searchpath)
        matchfiles = [os.path.join(searchpath, basename) for basename in allfiles if basename.endswith(searchtype) and searchfile in basename]
        try:
            foundfile = max(matchfiles, key=os.path.getctime)
        except ValueError:
            sys.exit(input('Could not find file containing "' + searchfile + '" in folder "' + searchpath + '". Ensure that you have downloaded the correct file, or adjust variable.json search parameters.'))
        return foundfile

    def find_excluded_pages(self):
        """Return list of all page numbers (array starting 0) that contain an exname"""
        pages = []
        # Open the pdf file
        reader = PdfReader(self.editfile)
        for i in range(0, len(reader.pages)):
            # Progress counter
            if i + 1 < len(reader.pages):
                progress = str(i + 1) + '/' + str(len(reader.pages))
            else:
                progress = 'Complete\n'
            print('Reading pages "' + self.editfile + '"... ' + progress, end='\r', flush=True)
            
            # Skip pages that have already been checked
            if i in pages:
                continue
            
            # Itterate through all names in self.exnames
            for name in self.exnames:
                
                # If name is on the page add it to the list of pages to skip in final outfile
                if re.search(name,reader.pages[i].extract_text()):
                    pages.append(i)
                    
                    # Check if the next page belongs to that person too
                    if i + 1 == len(reader.pages): # If there is no next page, break the loop
                        break
                    
                    # If the next page doesn't contain the words 'Total Hours' it is a second page for the same employee, add it to the excluded pages list
                    if not re.search('Total Hours', reader.pages[i+1].extract_text()):
                        pages.append(i+1)
                    
                    # Remove the name of the person from the excluded names list to save search time in next itteration
                    self.exnames.remove(name)
                    break
            
            continue
        return pages

    def create_rect(self):
        """Draw a rectangle on exery page in a pdf"""
        # Open the pdf
        try:
            doc = fitz.open(self.editfile)
            print('Removing footer... ', end='')
            for page in doc:
                # For every page, draw a rectangle on coordinates (1,1)(100,100)
                page.draw_rect([-5,-173,725,-173],  color = (1, 1, 1), width = 12)
            # Save pdf
            print('Complete')
            doc.save(self.temp)
        except FileNotFoundError:
            sys.exit(input('\nCould not find file "' + self.editfile + '. Ensure file exists in directory.'))

    def watermark_merge(self):
        """Merge the watermark file to each page of the temp file not on the excluded pages list and create outfile"""

        with open(self.temp, "rb") as input_file, open(self.watermark, "rb") as watermark_file:
            input_pdf = PdfReader(input_file)
            watermark_pdf = PdfReader(watermark_file)
            watermark_page = watermark_pdf.pages[0]
            output = PdfWriter()

            for i in range(0, len(input_pdf.pages)):
                # Progress counter
                if i + 1 < len(input_pdf.pages):
                    progress = str(i + 1) + '/' + str(len(input_pdf.pages))
                else:
                    progress = 'Complete\n'
                print('Adding signature lines... ' + progress, end='\r', flush=True)

                # Skip excluded pages
                if i in self.expages:
                    continue

                # Merge pages
                pdf_page = input_pdf.pages[i]
                pdf_page.merge_page(watermark_page)    
                output.add_page(pdf_page)
            # Write the merged file
        with open(self.outfile, "wb") as merged_file:
            output.write(merged_file)
 
    def _debug(self, localvars=None, scope=None, globalvars=False, **kwargs):
        """
        Print passed variables in terminal
        Useful for .exe scripts 
        """
        print('\n===== DEBUG MODE =====\nTo turn off debug mode, change "debug" : "active" \nvariable to false in "variables.json" file.')
        
        if globalvars:
            print('\n... GLOBALS ...')
            for key in globals():
                print(key + ' : ' + str(globals()[key]))
        
        if scope:
            print('\n... SCOPE ...')
            for v in scope:
                try:
                    val = eval(v)
                    print(v, " : ", end='', flush=True)
                    print('type=', type(val), ' : ', end='', flush=True)
                    print('value=', val)
                    continue
                except:
                    print(v, ' : type=', type(v))
                    pass
        
        if localvars:
            print('\n... LOCALS ...')
            for key in localvars:
                print(key + ' : ' + str(localvars[key]))
        
        if kwargs:
            # Any other variable values called in variables.json ['debug']['kwargs']
            print('\n... KWARGS ...')
            for key, value in kwargs.items():
                    print(key, '=', eval(value))

        input('\n--- PRESS ANY KEY TO CONTINUE ---')


if __name__ == '__main__':
    TS_Format('variables.json')