import http.client, urllib.request, urllib.parse, urllib.error, base64
import json
import csv

#Make sure to fill in the blanks. To get an API key you'll need to set up an account and generate it through the Microsoft dashboard.
BING_API_KEY = ''
output_filename = '.csv'

#The first criterion is the only one validated against results - should be the "must have", like a person's name.
#
#Make sure the two lists are of equal length, or you'll get errors. If you're looking up the same criterion, like a title/company/location, with multiple names, for example,
#use an Excel fill or something to get a long list of the same word/phrase.
criterion_one = []
criterion_two = []

output_rows = []

count = 0
total_items = len(criterion_two)

for crit_one, crit_two  in zip(criterion_one, criterion_two): # zip to loop over the criteria lists simultaneously
    
    query = f'"{crit_one}"' + ' ' + f'"{crit_two}"' + ' ' + 'site:linkedin.com'
    
    try:
        print(f'Item {str(count)} of {str(total)}')
        headers = {'Ocp-Apim-Subscription-Key': BING_API_KEY }
        params = urllib.parse.urlencode({'q': query,'count': '50'}) # returns top 2 (German) results
        conn = http.client.HTTPSConnection('api.cognitive.microsoft.com')
        conn.request("GET", "/bing/v7.0/search?%s" % params, "{body}", headers)
        response = conn.getresponse()
        data = response.read().decode('utf-8')
        json_file = json.loads(data)
        conn.close()

        if 'webPages' in json_file.keys():
            for result in json_file['webPages']['value']:
                page_title = result['name']
                print(f"\tProcessing result {page_title} from Item {str(count)}.")
                if crit_one.lower() in page_title.lower(): # checks if the first criteron appears in the page title
                    if (('linkedin.com/in/' in result['url']) or ('linkedin.com/pub/' in result['url'])): # checks if the search result URL is a LI profile
                        if 'searchTags' in result.keys():
                            output_rows.append([crit_one, crit_two, result['url'], result['name'], result['snippet'], result['searchTags']])
                        else:
                            output_rows.append([crit_one, crit_two, result['url'], result['name'], result['snippet'], ""])
                        break
        else:
            output_rows.append([crit_one, crit_two, "no webpage results", "", "", ""])
   
    except Exception as e:
        print("\t" + "Exception: " + e + " " + e.msg)
        #outputs a row with the error data instead of a blank
        output_rows.append([crit_one, crit_two, "Exception Error", e, e.msg])
        continue

    count += 1
# Writing the results to a CSV file
with open(output_filename, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    for row in output_rows:
        writer.writerow(row)
                
print('Approx number of LinkedIn profiles found:')
print(len(output_rows))
