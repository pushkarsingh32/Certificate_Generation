import cv2
import pandas as pd

# template_path = 'templates/template1-1.png'
# details_path = 'certificate_prospects/Interns Datails.xlsx'
# output_path = 'generated_certificates'

# The input file contains names as a comma seperated list of format (Rank, User handle, Name)
participants = pd.read_excel("certificate_prospects/Interns Datails copy.xlsx")

template_file_path = 'templates/template1-1.png'

# Make sure this output directory already exists or else certificates won't actually be generated
output_directory_path = 'generated_certificates'

font_size = 3.8
# Below if taking value in BGR format not RGB
font_color = (68, 102, 253)

# Test with different values for your particular Template
# This variables determine the exact position where your text will overlay on the template
# Y adjustment determines the px to position above the horizontal center of the template (may be positive or negative)
coordinate_y_adjustment = 78
# X adjustment determiens the px to position to the right of verticial center of the template (may be positive or negative)
coordinate_x_adjustment = 7

for index, row in participants.iterrows():
    # print(row["User handle"])
    certiName = row["Name"].title()
    from_date = row["from"]
    to_date = row["to"]
    starting_month = row["s_month"]
    ending_month = row["e_month"]
    # print(f"{from_date} , {to_date}, {starting_month}, {ending_month}")
    img = cv2.imread(template_file_path)
    font = cv2.FONT_HERSHEY_SIMPLEX
    #
    text = certiName
    textsize = cv2.getTextSize(text, font, font_size, 10)[0]
    text_x = (img.shape[1] - textsize[0]) / 2 + coordinate_x_adjustment
    text_y = (img.shape[0] + textsize[1]) / 2 - coordinate_y_adjustment
    text_x = int(text_x)
    text_y = int(text_y)
    #
    cv2.putText(img, text, (text_x, text_y), font, font_size, font_color, 10)

    certiPath = "{}/{}.png".format(output_directory_path, row["Name"])
    cv2.imwrite(certiPath, img)

cv2.destroyAllWindows()

# certi_path = output_directory_path + certiName + '.png'
# cv2.imwrite(certi_path, img)