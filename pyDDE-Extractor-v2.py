#!/usr/bin/python

'''
    Shitty Codes to extract DDE
    by Jacob Soo, @_jsoo_

    Hashes for samples:
    f945105f5a0bc8ea0d62a28ee62883ffc14377b6abec2d0841e88935fd8902d3
    09287128aaf96479f0aca8eedc3c78d3e863aae1368ed9eb62b5c2df98f92810
    58bc300c0ab90fbbb1f51482eaf83ac274429ba81ea581fc56c2ecb2b501bed0
    4654664436b505044cb9609595f3967bcbc8035e75904f81d8610c33268edc74
'''

__author__ = "Jacob Soo, @_jsoo_"
__version__ = "0.1"

import zipfile
import sys, os, re

def _log(szString):
    print(szString)

def ExtractDDE(szInputFile, szPath):
    paragraphs = []
    try:
        document = zipfile.ZipFile(szInputFile)
        xml_content = document.read(szPath)
        document.close()
        
        matchObj = re.findall(r'(\<w\:fldSimple w\:instr\=\"(.*?)\"\>)|(\<w\:instrText xml\:space\=\"preserve\"\>(.*?)\<\/w\:instrText\>)|(\<w\:instrText>(.*?)\<\/w\:instrText>)', xml_content, re.DOTALL|re.UNICODE)
        for item in matchObj:
            if '<w:instrText xml' in item[2]:
                paragraphs.append(item[3])
            elif '<w:instrText>' in item[4]:
                paragraphs.append(item[5])
            elif '<w:fldSimple w:instr=' in item[0]:
                paragraphs.append(item[1])
    except:
        _log("[-] File possibly doesn't exist in %s!" % szPath)
    return ''.join(paragraphs)
    

if __name__ == '__main__':
    if (len(sys.argv) < 2):
        _log("[+] Usage: %s [Path_To_DOCX]" % sys.argv[0])
        sys.exit(0)
    else:
        bFound = False
        szIdentifier = ""
        szPath = ""
        _log("[+] Finding DDE from %s" % sys.argv[1])
        
        with zipfile.ZipFile(sys.argv[1], 'r') as f:
            names = f.namelist()
            for name in names:
                if 'xl/externalLinks/' in name:
                    szIdentifier = "xlsx"
                    break
                elif 'word/document.xml' in name:
                    szIdentifier = "docx"
                    break
            f.close()
            
            if szIdentifier == 'xlsx':
                _log("[*] This is a xlsx sample.")
                locationArr = ['xl/externalLinks/externalLink1.xml']
                for iXML in locationArr:
                    try:
                        document = zipfile.ZipFile(sys.argv[1])
                        xml_content = document.read(iXML)
                        document.close()
                        matchObj = re.findall(r'ddeService\=[\'|\"](.*?)[\'|\"]', xml_content, re.DOTALL|re.UNICODE)
                        matchObj2 = re.findall(r'ddeTopic\=[\'|\"](.*?)[\'|\"]', xml_content, re.DOTALL|re.UNICODE)
                        
                        if matchObj and matchObj2:
                            _log("[+] Found the following DDE in : %s" % iXML)
                            _log("\t%s %s" % (matchObj[0], matchObj2[0]))
                            bFound = True
                            break
                    except:
                        _log("[-] Probably some issues with my codes")
            elif szIdentifier == 'docx':
                _log("[*] This is a docx sample.")
                locationArr = ['word/document.xml', 'word/header1.xml', 'word/header2.xml', 'word/footer1.xml', 'word/footer2.xml']
                for iXML in locationArr:
                    try:
                        DDE_Codes = ExtractDDE(sys.argv[1], iXML)
                        if DDE_Codes and not bFound:
                            _log("[+] Found the following DDE in : %s" % iXML)
                            _log("\t%s" % DDE_Codes)
                            bFound = True
                            break
                        #else:
                        #    _log("[-] Couldn't find anything at all :(")
                    except:
                        _log("[-] Probably some issues with my codes")