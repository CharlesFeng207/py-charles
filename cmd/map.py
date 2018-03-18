import sys
import json

if len(sys.argv) > 1:
    path = sys.argv[1]
    print path

    with open(path) as json_file:
        json_data = json.load(json_file)
        print(json_data["MsgId"])
else:
    print "parameters error!"