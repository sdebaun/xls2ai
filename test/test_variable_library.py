'''
Created on Dec 2, 2013

@author: sdebaun
'''
import unittest

from fixtures import MOCK_ARGS

from xls2ai import VariableLibrary

class TestVariableLibrary(unittest.TestCase):
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