library(XML)
##use R's inbuilt unzip function, knowing that the required metadata is in docProps/core.xml
doc = xmlInternalTreeParse(unzip('../../data/6332 A_rep_cijferlijst.xlsx','docProps/core.xml'))
##define the namespace
ns=c('dc'= 'http://purl.org/dc/elements/1.1/')
##extract the author using xpath query
author = xmlValue(getNodeSet(doc, '/*/dc:creator', namespaces=ns)[[1]])
