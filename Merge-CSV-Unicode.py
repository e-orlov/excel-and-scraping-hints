import os
import glob
import pandas as pd
os.chdir("C:/Users/evgeniy/Downloads/Project_ Lokale KWs ohne Stadtbezug Mobile (1)/domains")
extension = 'csv'
all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
combined_csv = pd.concat([pd.read_csv(f) for f in all_filenames ])
combined_csv.to_csv( "combined_csv.csv", index=False, encoding='utf-8-sig')