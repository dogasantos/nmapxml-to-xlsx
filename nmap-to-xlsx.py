#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = '0.1'

import argparse
import sys
import os
import nmap
import xlsxwriter

def nmap_LoadXmlObject(filename):
    nm = nmap.PortScanner()
    nxo = open(filename, "r")
    xmlres = nxo.read()
    nm.analyse_nmap_xml_scan(xmlres)
    return nm


def banner():
    print(" nmap-to-xlsx "+str(__version__)+" @ dogasantos")
    print("------------------------------------------------")
    print(" convert nmap xml report file into a xlsx file")
    print("------------------------------------------------")

def parser_error(errmsg):
    banner()
    print("Usage: python " + sys.argv[0] + " [Options] use -h for help")
    print("Error: %s" %errmsg)
    sys.exit(1)

def parse_args():
    parser = argparse.ArgumentParser(epilog='\tExample: \r\npython ' + sys.argv[0] + "[options]")
    parser.error = parser_error
    parser._optionals.title = "Options:"
    parser.add_argument('-x', '--xml', help="Nmap xml report file", required=True)
    parser.add_argument('-o', '--xlsx', help="xlsx output file", required=True)
    return parser.parse_args()


def nmap_GetScriptData(nmapObj, host, proto, port, scriptname):
    try:
        data=nmapObj[host][proto][int(port)]['script'][scriptname]
    except:
        data=False
    return data




def Convert(nmapObj,xlsxfilename):
    workbook = xlsxwriter.Workbook(xlsxfilename)
    worksheet = workbook.add_worksheet()

    cell_format_t = workbook.add_format()
    cell_format_t.set_bg_color('#082099')
    cell_format_t.set_bold()
    cell_format_t.set_font_color('#FFFFFF')

    cell_format_1 = workbook.add_format()
    cell_format_1.set_bg_color('#c9e6ff')

    cell_format_2 = workbook.add_format()
    cell_format_2.set_bg_color('#b0daff')


    worksheet.write(0,0, "IP ADDRESS", cell_format_t)
    worksheet.write(0,1, "PORT", cell_format_t)
    worksheet.write(0,2, "SERVICE", cell_format_t)
    worksheet.write(0,3, "SERVICE DETAILS", cell_format_t)
    worksheet.write(0,4, "SCRIPT NAME", cell_format_t)
    worksheet.write(0,5, "SCRIPT OUTPUT", cell_format_t)

    row = 1
    col = 0
    script_nro = 0
    count = 0
    for ip in nmapObj.all_hosts():
        print("[*] Host: " + ip)
        openports = nmapObj[ip]['tcp'].keys()

        if count % 2 == 0:
            cf = cell_format_1
        else:
            cf = cell_format_2
        count += 1

        for port in openports:
            service_details = nmapObj[ip]['tcp'][port]
            worksheet.write(row, 0, ip, cf)
            worksheet.write(row, 1, port, cf)
            worksheet.write(row, 2, service_details['name'], cf)
            worksheet.write(row, 3, service_details['product'] + " " + service_details['version'], cf)
            try:
                scripts = service_details['script']
                for script_name in scripts:
                    if script_nro == 0:
                        worksheet.write(row, 4, script_name, cf)
                        worksheet.write(row, 5, service_details['script'][script_name].replace("\\n", "\n"), cf)
                    else:
                        worksheet.write(row, 0, ip, cf) 
                        worksheet.write(row, 1, port, cf)
                        worksheet.write(row, 2, service_details['name'], cf)
                        worksheet.write(row, 3, service_details['product'] + " " + service_details['version'], cf)
                        worksheet.write(row, 4, script_name, cf)
                        worksheet.write(row, 5, service_details['script'][script_name].replace("\\n", "\n"), cf)
                    script_nro += 1
                    row += 1
            except:
                pass
            row += 1
    workbook.close()
    return True

if __name__ == "__main__":

    projectname = "default"
    args = parse_args()

    xml_report = args.xml
    xlsx_output = args.xlsx

    nmapObj = nmap_LoadXmlObject(xml_report)
    Convert(nmapObj,xlsx_output)
    
