import base64
import json

def base64_txt_to_json(input_file, output_file=None):

    with open(input_file, "r") as file:
        base64_str = file.read().strip()

    decoded_str = base64.b64decode(base64_str).decode("utf-8")


    json_data = json.loads(decoded_str)


    if output_file:
        with open(output_file, "w") as outfile:
            json.dump(json_data, outfile, indent=4)

    return json_data


data = base64_txt_to_json("json_b64.txt","op.json")
print(data)


