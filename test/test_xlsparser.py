'''
Created on Dec 2, 2013

@author: sdebaun
'''
import unittest
from mock import patch

from fixtures import MOCK_ARGS, TEST_XLS

from xls2ai import VariableLibrary, XLSParser

@patch('xls2ai.VariableLibrary',autospec=True)
@patch('xls2ai.FileWriter',autospec=True)
class TestXLSParser(unittest.TestCase):
    
    def setUp(self):
        self.sut = XLSParser(TEST_XLS)
        
    def test(self, _FileWriter, _VariableLibrary):
        self.sut.parse()
        _VariableLibrary.assert_called_once_with(MOCK_ARGS)
        _FileWriter.assert_called_once_with('testcards',renderer=_VariableLibrary.return_value)
        _FileWriter.return_value.write.assert_called_once_with()

if __name__ == "__main__":
    #import sys;sys.argv = ['', 'Test.testName']
    unittest.main()