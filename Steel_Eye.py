import xlrd
import json
import boto3

#-----------------------------------------------------------------
## Creating the .json file as specified

# Opening the workbook 
wb = xlrd.open_workbook('/home/yeswanth/Downloads/ISO_MIC.xls') # please change the directory accordingly

# Reading the tab of MICs List by CC
MIC_CC = wb.sheet_by_name('MICs List by CC')

# Assigning the first row as keys
keys = []
for col_index in range(MIC_CC.ncols):
	keys.append(MIC_CC.cell_value(0,col_index))

# list to hold dictionaries
dict_list = []

# Create a list of dict containing all rows (except row 1).
# The values in row 1 would be the keys for in each dict.
for row_index in range(1, MIC_CC.nrows):
	dict = {}
	for col_index in range(MIC_CC.ncols):
		dict[keys[col_index]] = MIC_CC.cell_value(row_index, col_index)
	dict_list.append(dict)
 
# Serialize the list of dicts to JSON as a string
# j is a json string now
j = json.dumps(dict_list)

# Write to file to data.json
with open('data.json', 'w') as f:
    f.write(j)

# ----------------------------------------------------------------------
## Storing the above created .json file in AWS S3 bucket

# Connect to S3 with access credentials 
# please put your credentials
session = boto3.Session(
    aws_access_key_id=AWS_SERVER_PUBLIC_KEY,
    aws_secret_access_key=AWS_SERVER_SECRET_KEY,
)
s3 = session.resource('s3')

# Creating a bucket with name as 'mybucket'
s3.create_bucket(Bucket='mybucket')
s3.create_bucket(Bucket='mybucket', CreateBucketConfiguration={'LocationConstraint':'my_location'}) # please put your location instead of 'my_location'

# Storing the .json file in bucket
s3.Object('mybucket', 'data.json').put(Body=open('/home/yeswanth/data.json', 'rb')) # please change the directory of file .json accordingly

