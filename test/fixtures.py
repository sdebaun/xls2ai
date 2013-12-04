TEST_XLS = 'testsource.xls'

MOCK_ARGS = {'variables': [{'varName':'foo', 'textcontent':'true' },
                           {'varName':'bar', 'visibility':'true' },
                           {'varName':'bar2', 'visibility':'true' },
                           {'varName':'baz', 'fileref':'true' },
                           {'varName':'dsid', 'textcontent':'true' }],
             'datasets': [{'dataSetName':'01', 'vars': [{'varName':'foo','value':'<p>Some Thing</p>'},
                                                {'varName':'bar','value':'false'},
                                                {'varName':'bar2','value':'false'},
                                                {'varName':'baz','value':'file:///something.png'},
                                                {'varName':'dsid','value':'<p>01</p>'}]},
                          {'dataSetName':'02', 'vars': [{'varName':'foo','value':'<p>Other Thing</p>'},
                                                {'varName':'bar','value':'true'},
                                                {'varName':'bar2','value':'true'},
                                                {'varName':'baz','value':'file:///otherthing.png'},
                                                {'varName':'dsid','value':'<p>02</p>'}]},
                          ]
             }

