with open ('temp.csv', 'wb') as outfile:
    outfile.writelines(data_in[0])
    outfile.writelines(data_in[5:])