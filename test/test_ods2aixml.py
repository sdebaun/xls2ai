'''
Created on Dec 2, 2013

@author: sdebaun
'''
import unittest
from mock import MagicMock, patch

from ods2aixml import VariableLibrary, XLSParser

MOCK_ARGS = {'variables': [{'varName':'foo', 'textcontent':'true' },
                           {'varName':'bar', 'visibility':'true' },
                           {'varName':'bar2', 'visibility':'true' },
                           {'varName':'baz', 'fileref':'true' }],
             'datasets': [{'dataSetName':'01', 'vars': [{'varName':'foo','value':'<p>Some Thing</p>'},
                                                {'varName':'bar','value':'false'},
                                                {'varName':'bar2','value':'false'},
                                                {'varName':'baz','value':'file:///something.png'}]},
                          {'dataSetName':'02', 'vars': [{'varName':'foo','value':'<p>Other Thing</p>'},
                                                {'varName':'bar','value':'true'},
                                                {'varName':'bar2','value':'true'},
                                                {'varName':'baz','value':'file:///otherthing.png'}]},
                          ]
             }

@patch('ods2aixml.VariableLibrary',autospec=True)
@patch('ods2aixml.FileWriter',autospec=True)
class TestXLSParser(unittest.TestCase):
    
    def setUp(self):
        self.sut = XLSParser('testsource.xls')
#         self.sut = ODSParser('100kw-components.ods')
        
    def test(self, _FileWriter, _VariableLibrary):
        self.sut.parse()
        _VariableLibrary.assert_called_once_with(MOCK_ARGS)
        _FileWriter.assert_called_once_with('testcards',renderer=_VariableLibrary.return_value)
        _FileWriter.return_value.write.assert_called_once_with()
        


# @patch('ods2aixml.VariableLibrary',autospec=True)
# @patch('ods2aixml.FileWriter',autospec=True)
# class TestODSParser(unittest.TestCase):
#     
#     def setUp(self):
#         self.sut = ODSParser('testsource.ods')
# #         self.sut = ODSParser('100kw-components.ods')
#         
#     def test(self, _FileWriter, _VariableLibrary):
#         self.sut.parse()
#         _VariableLibrary.assert_called_once_with(MOCK_ARGS)
#         _FileWriter.assert_called_once_with('testcards',renderer=_VariableLibrary.return_value)
#         _FileWriter.return_value.write.assert_called_once_with()
        

class TestRenderer(unittest.TestCase):
    def setUp(self):
        self.sut = VariableLibrary(MOCK_ARGS)
        self.out = self.sut.render()
         
    def testOutputContainsVariables(self):
        self.assertIn('<variable  category="&ns_flows;" trait="textcontent" varName="foo"></variable>', self.out)
        self.assertIn('<variable  category="&ns_vars;" trait="visibility" varName="bar"></variable>', self.out)
        self.assertIn('<variable  category="&ns_vars;" trait="fileref" varName="baz"></variable>', self.out)
     
    def testOutputContainsDatasets(self):
        self.assertIn('<foo><p>Some Thing</p></foo>', self.out)
        self.assertIn('<foo><p>Other Thing</p></foo>', self.out)
        self.assertIn('<bar>true</bar>', self.out)
        self.assertIn('<bar>false</bar>', self.out)
        self.assertIn('<baz>file:///something.png</baz>', self.out)
        self.assertIn('<baz>file:///otherthing.png</baz>', self.out)
    
if __name__ == "__main__":
    #import sys;sys.argv = ['', 'Test.testName']
    unittest.main()