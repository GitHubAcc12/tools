import argparse
import pandas as pd
import os
from pathlib import Path

if __name__=='__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('-i', '--input', default='.',
                    required=False, help='path to folder with excel files')
    ap.add_argument('-o', '--output', default='/csv', required=False, help='where to save the csv files')
    args = vars(ap.parse_args())

    i_directory = os.fsencode(args['input'])

    o_directory = os.fsencode(args['input']+args['output'])
    Path(args['output']).mkdir(exist_ok=True)


    for file in os.listdir(i_directory):
        filename = os.fsdecode(file)
        if filename.endswith(".xls") or filename.endswith(".xlsx"): 
            try:
                data = pd.read_excel(args['input']+'/'+filename, index_col=0)
                with open(args['output']+'/'+filename[:-filename.find('.xls')+1] + '.csv', 'w') as csvfile:
                    data.to_csv(csvfile)
            except KeyboardInterrupt:
                raise
            except:
                print(f'File {filename} failed to convert, skipping')
        