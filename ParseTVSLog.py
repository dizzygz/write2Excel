#!/usr/bin/python

import sys, getopt
import os
import re,time
import xlwt

def output(filename, sheet,dataFile):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)
    inputfile = dataFile

    variables = ["Date", "Time", "Vid", "m_logCodes","m_strCmds","m_dir","m_entryDir","m_lastStopIdAccepted","GetVelocity","m_tp","m_bTpbCommanded","m_tpBias","m_finePosition",
 "m_fpm","m_numVirtualXover", "m_bAccurateTagAccepted", "m_distSinceLastTag", "m_distanceSinceLastXover", "m_rawTp", "m_vccNum", "m_loopNum", "m_rawCmdDir", "m_lastPollTime","m_lid",
 "m_segmentID", "m_trackNum", "m_brakeRate", "m_cmdDir", "m_cmdTxidValue","m_coarseOffset","m_coarsePos","m_commLossFlg", "m_crossoverNum", "m_currTagDir","m_ebcApplyEB","m_refCrossoverNum",
 "m_referencedPos","m_referencedTagDir","m_refPosition","m_refTagPosition","m_reportPosition","m_resetCommFlg", "m_simXovers","m_storageModeStatus","m_tagOffsetDist",
 "m_tagOffsetPos","m_trainAtStationStop","m_vTargetAchieved","m_xoverDetectionDistance","m_ascState","m_bCsdeEnabled","m_bDocked","m_bHasEAO","m_bJustWentActive","m_referencedTag",
 "m_currTag","m_antPt", "Bits"]

    icol=0
    for icol,col in enumerate(variables):
       sh.write(0, icol, col)
       #++icol

    x_desc = 'Display'
    y_desc = 'Dominance'
    z_desc = 'Test'
    desc = [x_desc, y_desc, z_desc]

    f = open(inputfile, mode='r', buffering=-1, encoding='utf-16', errors='strict',
             newline=None, closefd=True, opener=os.open)

    n=1
    for line in f:
         if re.match(r"\d{4}.\d{2}.\d{2}\s+\d{2}\:\d{2}\:\d{2}\.\d{3}\s+V[1-9]*\s+[0-1]", line):
             line = line.lstrip(' ')
             data = re.split("\s+",line)
             for icol2,v in enumerate(data):
                sh.write(n, icol2, v)

             n+=1
    book.save(filename)


def main(argv):
   inputfile = ''
   outputfile = ''
   try:
      opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
   except getopt.GetoptError:
      print ('test.py -i <inputfile> -o <outputfile>')
      sys.exit(2)
   for opt, arg in opts:
      if opt == '-h':
         print ('test.py -i <inputfile> -o <outputfile>')
         sys.exit()
      elif opt in ("-i", "--ifile"):
         inputfile = arg
      elif opt in ("-o", "--ofile"):
         outputfile = arg
   t = time.time()
   output("Result"+str(t)+".xls","TVSLog",inputfile)

if __name__ == "__main__":
   main(sys.argv[1:])



