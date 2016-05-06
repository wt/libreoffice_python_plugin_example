#!/bin/bash

make
unopkg remove example-wt.oxt
unopkg add example-wt.oxt
libreoffice
