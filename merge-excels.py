import glob
import pandas as pd
from datetime import datetime
from timeit import default_timer as timer

from pandas.core.indexes.base import Index

def main():
    now = datetime.now()
    output_file_name = 'merged_' + str(now.strftime('%Y-%m-%d_%H-%M-%S')) + '.xlsx'
    
    all_data = []
    total_rows = 0

    for file in glob.glob('./in/*.xlsx'):
        df = pd.read_excel(file, header=None)
        rows = len(df.index)
        total_rows = total_rows + rows
        print(f'File: {file} dimensions: {df.shape[0]} rows, {df.shape[1]} columns')
        all_data.append(df)

    print(f'Generating output file (./out/{output_file_name})...')
    start = timer()

    result = pd.concat(all_data, ignore_index=False, join='inner', axis=0)
    result.to_excel(f'./out/{output_file_name}', index=False, header=None)

    end = timer()
    print(f'DONE! Output file dimensions: {result.shape[0]} rows, {result.shape[1]} columns. Merging process took {round(end-start, 2)} seconds')

if __name__ == '__main__':
  main()