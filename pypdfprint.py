#imports
from natsort import natsorted
import win32api
import win32print
import yaml
import pathlib
import tempfile
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
import re



#class
class pypdfprint:
    def __init__(self,
                 file : str,
                 printer = None,
                 pages = "all",
                 copy = 1 ,
                 collate = True,
                 print_order = 0
                ) :
        
        """
        file -> Document path 

        printer -> Select which printer to print to [if None; preset default printer is used]

        pages -> Which pages to be printed eg.-('all','1','1-5,12','odd','even') [Default:'all']

        copy -> Number of copies required [Default: 1]

        collate -> Whether collate the pages or not [Default: True]

        print_order -> Whether to print normally or in reverse [Default: 0 (First page first) ] (In reverse(1): first page is printed last, such that it can be taken directly without having to rearrange the pages.)
        """
        
        self.default_printer = self.settingsload()['defprinter']
        if printer : 
            self.default_printer = printer
        win32print.SetDefaultPrinter(self.default_printer)

        param_dict = {
            'file' : file,
            'printer' : printer, 
            'pages' : pages, 
            'copy' : copy, 
            'collate' : collate, 
            'print_order' : print_order, 
        }

        self.errorclosures(param_dict)
        
        tempdir = self.tempdirpdf(file)
        param_dict['tempdir'] = tempdir
 
        pagelist = self.pagelist(parameters=param_dict)

        printlist = self.printlist(pagelist=pagelist,parameters=param_dict)

        for i in printlist:
            self.sendprint(i)

        os.rmdir(tempdir)

        pass
    

    def errorclosures(self,parameters: dict):
        if pathlib.Path(parameters['file']).suffix != ".pdf":
            raise ValueError("File not a pdf")

        if parameters['printer'] == None and self.default_printer not in self.list_printers():
            # If printer was not given as param and default printer from yaml is not found 
            raise ValueError(f"Default printer({self.default_printer}) not found.\nProvide printer name in param or change default printer\nAvailable printers: {self.list_printers()}")
            
        if parameters['printer'] != None and parameters['printer'] not in self.list_printers():
            # If printer was given as param but not found as available printer
            raise ValueError("printer name in param not found. Verify printer name using list_printers function")
        
        if parameters['print_order'] not in [0,1]:
            raise ValueError("Impossible event for print_order")
        
        pagenum_pattern = '\d{1,}((-\d{1,})|(,\d{1,})?){1,}'
        if not re.fullmatch(pagenum_pattern,parameters['pages']) and parameters['pages'] not in ['all','odd','even']:
            raise ValueError("Impossible event for pages")
        
        if type (parameters['copy']) != int:
            raise ValueError("Impossible event for copy")

        if parameters['copy'] < 1 :
            raise ValueError("Impossible event for copy")
        

        return

    def settingsload(self) -> dict:
        file = open('settings.yaml', 'r')
        settings = yaml.load(file,Loader=yaml.FullLoader)
        return settings

    def list_printers(self) -> list:
        all_printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        return all_printers
    
    def set_printer(self,printer_name):
        data_dict = {'defrinter':printer_name}
        
        with open('settings.yaml', 'w') as outfile:
            yaml.dump(data_dict, outfile, default_flow_style=False)
            self.default_printer = printer_name
            win32print.SetDefaultPrinter(self.default_printer)

        return

    def tempdirpdf(self,fp) -> str:
        inputpdf = PdfFileReader(open(fp, "rb"))
        tempdir = tempfile.mkdtemp()
        
        for i in range(inputpdf.numPages):
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            path = os.path.join(tempdir,"document-page%s.pdf" % str(i+1))
            with open(path,"wb") as outputStream:
                output.write(outputStream)### functions to add individual page pdf files into the folder 
        return tempdir

    def printlist(self,parameters: dict,pagelist: list,) -> list:
        ref_dict = {}
        tempdir = parameters['tempdir']
        count = 1

        for i in natsorted(os.listdir(tempdir)):
            ref_dict[count] = os.path.join(tempdir,i)
            count += 1
        
        arr = []
        for i in pagelist: arr.append(ref_dict[i])

        if parameters['copy'] > 1:
            if parameters['collate'] == True:
                arr = arr*parameters['copy']
                pass

            elif parameters['collate'] == False:
                temparr = []
                for i in arr:
                    for j in range(parameters['copy']):
                        temparr.append(i)
                arr = temparr
                pass
            
        
        if parameters['print_order'] == 1: arr = arr[::-1]

        finalarr = arr
        return finalarr
        
    def pagelist(self,parameters : dict,) -> list:
        """Returns an array containing all the pages that needs to be printed"""
        files = os.listdir(parameters['tempdir'])
        noofpages = len(files)


        param_pages = parameters['pages']

        arr = []
        if param_pages == "all":
            for i in range(noofpages) : arr.append(i+1)
            return arr

        elif param_pages == "odd":
            for i in range(0,noofpages,2) : arr.append(i+1)
            return arr

        elif param_pages == "even":
            for i in range(1,noofpages,2) : arr.append(i+1)
            return arr

        else:
            split = param_pages.split(',')
            arr = []

            for i in split:
                if i.isnumeric():
                    arr.append(int(i))

                elif "-" in i:
                    scndsplit = i.split("-")
                    first = int(scndsplit[0])
                    second = int(scndsplit[1])
                    if second < first: raise ValueError("Imopssible event for pages")

                    for i in range(first,second+1): arr.append(i)

        arr = list(set(arr))   
        for i in arr: 
            if int(i) > noofpages : raise ValueError("Impossible event for pages")
        
        arr = natsorted(arr)

        return arr
    
    def sendprint(self,file):
        win32api.ShellExecute(0, "print", file, None,  ".",  0)
        return


#testing
if __name__ == "__main__":
    testfile1 = "test/t1.pdf"
    testfile2 = "test/t2.pdf"

    x = pypdfprint(file = testfile1,copy = 1,pages='1,2',collate=True,print_order=0,printer='EPSON5BA3A3 (L3150 Series)')