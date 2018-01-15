# -*- coding:utf-8 -*-

from optparse import OptionParser
from XlsFileUtil import XlsFileUtil
from Log import Log
import os

def addParser():
    parser = OptionParser()

    parser.add_option("-f", "--filePath",
                      help="original.xls File Path.",
                      metavar="filePath")

    parser.add_option("-t", "--targetFloderPath",
                      help="Target Floder Path.",
                      metavar="targetFloderPath")

    parser.add_option("-i", "--iOSAdditional",
                      help="iOS additional info.",
                      metavar = "iOSAdditional")

    parser.add_option("-a", "--androidAdditional",
                      help="android additional info.",
                      metavar="androidAdditional")

    (options, args) = parser.parse_args()
    Log.info("options: %s, args: %s" % (options, args))

    return options


def startConvert(options):
    filePath = options.filePath
    targetFloderPath = options.targetFloderPath
    iOSAdditional = options.iOSAdditional
    androidAdditional = options.androidAdditional

    if filePath is not None:
        if targetFloderPath is None:
            Log.error("targetFloderPath is None！use -h for help.")
            return

        # xls
        Log.info("read xls file from"+filePath)
        xlsFileUtil = XlsFileUtil(filePath)

        table = xlsFileUtil.getTableByIndex(0)
        convertiOSAndAndroidFile(table,targetFloderPath,iOSAdditional,androidAdditional)

        Log.info("Finished,go to see it -> "+targetFloderPath)

    else:
        Log.error("file path is None！use -h for help.")


def convertiOSAndAndroidFile(table,targetFloderPath,iOSAdditional,androidAdditional):
    firstRow = table.row_values(0)# 第0行所有的值

    keys = table.col_values(0)# 第0列所有的值

    Log.info("targetFloderPath: " + targetFloderPath)

    if not os.path.exists(targetFloderPath):
            os.makedirs(targetFloderPath)


    fo = open(targetFloderPath+"/output.txt", "wb")
    Log.info("open file " + targetFloderPath+"/output.txt")

    for x in range(len(keys)):
            row = table.row_values(x)# 第0行所有的值

            if row[0] is None or row[0] == '' or row[1] is None or row[1] == '':
                Log.error("Key:" + keys[x] + "\'s value is None. Index:" + str(x + 1))
                continue
            content =  row[0] + "  " + row[1] + "\n"
            Log.info("wcontent is " + content)
            fo.write(content);


    Log.info("will close")
    fo.close()


def main():
    options = addParser()
    startConvert(options)

main()