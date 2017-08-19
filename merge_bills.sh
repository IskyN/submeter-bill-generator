#!/bin/bash

if [ "`pwd`" != "/mnt/c/Users/nisky/Desktop/Pact & Comfort Care Stuff/Pact Submeter Bills/submeter-bill-generator" ]; then
    echo "Current directory is not 'submeter-bill-generator'"
    exit 1
fi

if [ $# -eq 1 ]; then
    bill_folders=$1
else
    bill_folders=`find Bills -mindepth 1 -maxdepth 1 -type d -printf '%f '`
fi
# echo $bill_folders
for d in $bill_folders; do
    echo $d
    if [ -e 'Bills/'$d'_bills.pdf' ]; then continue; fi  # don't redo!

    find "Bills/"$d -mindepth 1 -maxdepth 1 -name '*.pdf' | xargs \
        gs -dBATCH -dNOPAUSE -sDEVICE=pdfwrite \
           -sOutputFile='Bills/'$d'_bills.pdf'
done

