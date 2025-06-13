import requests

url = "https://lendlease.service-now.com/sn_hr_core_case.do?PDF&sys_id=00016018db51b450ee773313e29619cc&sysparm_view=Default%20view"

headers = {
    'Authorization': 'Bearer FALrqmQzoUFhfStHntyhax7dXNmfOJ8YMNkX89VjLtDRhM5bnr4GdOeGHVpZJSFWF38JIRPbgCbF-yvWItegww',
    'Cookie': 'BIGipServerpool_lendlease=d0eb094de8ff43f5cc2be299bfd8eec0; JSESSIONID=50F9C7CAD0CC372494A46F314469BADD; glide_node_id_for_js=efe2bd7e650020ffb1a3817795c1b2947f8fae623442f4932f54498ae9e12757; glide_user_route=glide.f918fa4ee5b30c271260985b8b448be6'
}

response = requests.get(url, headers=headers)

# Check if the response is successful
if response.status_code == 200:
    # Save the PDF content to a file
    with open("test_output.pdf", "wb") as f:
        f.write(response.content)
    print("PDF downloaded successfully as 'test_output.pdf'.")
else:
    print(f"Failed to download PDF. Status code: {response.status_code}")
    print(response.text)
