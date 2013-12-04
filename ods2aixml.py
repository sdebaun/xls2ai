'''
Created on Dec 2, 2013

@author: sdebaun

TODO: check that XML is well formed after output
    see http://code.activestate.com/recipes/52256-check-xml-well-formedness/
    
TODO: Better: auto-escape contents, and add line break special character translation

TODO: STDOUT results

TODO: user-friendly error messages

'''

TEMPLATE = 'aixml-template.xml'

import sys, argparse, pystache, xlrd

class XLSParser(object):
    def __init__(self,filename):
        self.filename = filename

    def parse(self):
        xls = xlrd.open_workbook(self.filename)
        for sheet in xls.sheets():
            if 'EXPORT' in sheet.name:
                print 'Parsing ', sheet.name
                self._parse_sheet(sheet.name.split('EXPORT')[0].strip(), sheet)
                
    def _find_header_row(self, sheet):
        for index in xrange(0, sheet.nrows):
            if sheet.cell(index,0).value == 'DataSets':
                return index
        raise Exception("No 'DataSets' cell found")
        
    def _parse_sheet(self,name,sheet):
        variables, datasets = [None, ], []
        
        header_row = self._find_header_row(sheet)
            
        for index in xrange(1, sheet.ncols):
            vartype, varname = sheet.cell(header_row, index).value, sheet.cell(header_row+1, index).value
            print vartype, varname
            if vartype in ['textcontent', 'visibility', 'fileref']:
                variables.append({'varName':varname, vartype:'true'})
            else:
                variables.append(None)
        variables.append({'varName':'dsid', 'textcontent':'true'})
        print variables

        for index in xrange(header_row+2, sheet.nrows):
            dsn = sheet.cell(index, 0).value
            if dsn:
                dataset = {'dataSetName': '%02d' % int(dsn), 'vars':[]}
                for colindex in xrange(1, sheet.ncols):
                    vardef = variables[colindex]
                    if vardef:
                        content = sheet.cell(index, colindex).value
                        if 'textcontent' in vardef.keys(): content = "<p>%s</p>" % content
                        if 'visibility' in vardef.keys(): content = content and 'true' or 'false'
                        dataset['vars'].append({'varName':vardef['varName'],'value':content})
                dataset['vars'].append({'varName':'dsid','value':'<p>%02d</p>' % int(dsn)})
                datasets.append(dataset)
                print dataset
                
        vl = VariableLibrary({'variables':[v for v in variables if v], 'datasets':datasets}) 
        fw = FileWriter(name, renderer=vl)
        fw.write()

    def cellvalue(self,cell,p=False):
        try:
            valtype = cell.attrib['{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value-type']
        except:
            return None
        if valtype=='string': val = cell[0].text
        elif valtype=='float': val = cell.attrib['{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value']
        elif valtype=='boolean': val = (cell.attrib['{urn:oasis:names:tc:opendocument:xmlns:office:1.0}boolean-value']=='true') and 'true' or 'false'
        return p and ("<p>%s</p>" % val) or val
    

from ooopy.OOoPy import OOoPy

class ODSParser(object):
    def __init__(self,filename):
        self.filename = filename

    def parse(self):
        ods = OOoPy(self.filename)
        f = ods.read('content.xml')
        sheets = f.tree.getroot().find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body').find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}spreadsheet')
        for s in sheets:
            name = s.attrib['{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name']
            if 'EXPORT' in name:
                self._parse_sheet(name.split('EXPORT')[0].strip(), s)
    
    def cellvalue(self,cell,p=False):
        try:
            valtype = cell.attrib['{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value-type']
        except:
            return None
        if valtype=='string': val = cell[0].text
        elif valtype=='float': val = cell.attrib['{urn:oasis:names:tc:opendocument:xmlns:office:1.0}value']
        elif valtype=='boolean': val = (cell.attrib['{urn:oasis:names:tc:opendocument:xmlns:office:1.0}boolean-value']=='true') and 'true' or 'false'
        return p and ("<p>%s</p>" % val) or val
    
    def _parse_sheet(self,name,sheet):
        variables, datasets = [], []
        rows = sheet.findall('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-row')
        cols = sheet.findall('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-column')

#         for index, col in enumerate(cols[1:]):
#             varname, vartype = self.cellvalue(cols[index+1][1]), self.cellvalue(cell)

        for index, cell in enumerate(rows[0][1:]):
            varname, vartype = self.cellvalue(rows[1][index+1]), self.cellvalue(cell)
            print cell.tag, cell.attrib, cell.text
            if vartype in ['textcontent', 'visibility', 'fileref']:
                variables.append({'varName':varname, vartype:'true'})
            else:
                variables.append(None)
        print variables
        
        for row in rows[2:]:
            dsn = self.cellvalue(row[0])
            if dsn:
                dataset = {'dataSetName': '%02d' % int(self.cellvalue(row[0])),
                           'vars':[]}
                for index, cell in enumerate(row[1:]):
                    vardef = variables[index]
                    if vardef:
                        dataset['vars'].append({'varName':vardef['varName'],
                                                'value':self.cellvalue(cell,p='textcontent' in vardef.keys())})
                datasets.append(dataset)
                print dataset
                
        vl = VariableLibrary({'variables':[v for v in variables if v], 'datasets':datasets}) 
        fw = FileWriter(name, renderer=vl)
        fw.write()

class FileWriter(object):
    def __init__(self,name,renderer):
        self.name, self.renderer = name, renderer
        
    def write(self):
        open("%s.xml" % self.name, 'w').write(self.renderer.render())
    
class VariableLibrary(object):
    def __init__(self,args={}):
        for arg, item in args.iteritems():
            setattr(self,arg,item)

    def render(self):
        return pystache.Renderer().render(self)

class Converter(object):
    def __init__(self,infile):
        self.infile = infile
    
    def write(self):
        print(self)
        

if __name__=="__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('filename')
    XLSParser(**vars(parser.parse_args())).parse()
