#!/usr/bin/env python
# coding: utf-8

import pandas as pd
f=pd.read_csv("output_html_ap.csv")
keep_col = ['address','status code','status','Sim Score','Sim Match']
new_f = f[keep_col]
new_f.to_csv("newFile.csv", index=False)