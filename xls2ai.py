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

import argparse, pystache, xlrd

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
        raise Exception("No 'DataSets' cell found.")
        
    def _parse_sheet(self,name,sheet):
        header_row = self._find_header_row(sheet)
        
        variables = self._parse_variables(header_row, sheet)
        datasets = self._parse_datasets(header_row + 2, variables, sheet)
        
        vl = VariableLibrary({'variables':[v for v in variables if v], 'datasets':datasets}) 
        fw = FileWriter(name, renderer=vl)
        fw.write()
    
    def _parse_variables(self,header_row,sheet):
        variables = [None, ]
        for index in xrange(1, sheet.ncols):
            vartype, varname = sheet.cell(header_row, index).value, sheet.cell(header_row+1, index).value
            print vartype, varname
            if vartype in ['textcontent', 'visibility', 'fileref']:
                variables.append({'varName':varname, vartype:'true'})
            else:
                variables.append(None)
        variables.append({'varName':'dsid', 'textcontent':'true'})
        return variables
    
    def _parse_datasets(self,datasets_first_row, variables, sheet):
        datasets = []
        for index in xrange(datasets_first_row, sheet.nrows):
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
        return datasets
                
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
        return pystache.Renderer().render(self) # derives .mustache file from class name

class Converter(object):
    def __init__(self,infile):
        self.infile = infile
    
    def write(self):
        print(self)
        

if __name__=="__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('filename')
    XLSParser(**vars(parser.parse_args())).parse()
