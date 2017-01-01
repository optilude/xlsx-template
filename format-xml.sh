#!/bin/bash

# Helper script to extract test files to make them easier to visually inspect

mkdir extract

for x in *.xlsx ; do

    rm -r extract/$x
    mkdir extract/$x
    cd extract/$x
    unzip ../../$x

    for f in $(find . -name '*.xml') ; do

        mv $f $f.old
        xmllint --format $f.old > $f
        rm $f.old

    done

    cd ../..

done
