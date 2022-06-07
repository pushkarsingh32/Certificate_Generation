import cv2
import pandas as pd
from PIL import ImageFont, Image
import img2pdf


participants = pd.read_excel("certificate_prospects/health_interns.xlsx")
template_file_path = 'templates/template1.png'
output_directory_path = 'generated_certificates/'

font_color = (68, 102, 253)


def draw_text(img, to_draw_text, font, font_size, coordinate_x_adjust, coordinate_y_adjust):
    textsize = cv2.getTextSize(to_draw_text, font, font_size, 10)[0]
    text_x = (img.shape[1] - textsize[0]) / 2 + coordinate_x_adjust
    text_y = (img.shape[0] + textsize[1]) / 2 - coordinate_y_adjust
    cv2.putText(img, to_draw_text, (int(text_x), int(text_y)), font,
                font_size, font_color, 10)


def main_function():
    for index, row in participants.iterrows():
        img = cv2.imread(template_file_path)

        # For Intern Name
        intern_x_adjust, intern_y_adjust = 100, 78
        # intern_x_adjust, intern_y_adjust = 50, 78
        intern_name_text = row["Name"].title()
        font_size = 3.8
        intern_font = cv2.FONT_HERSHEY_SIMPLEX
        draw_text(img, intern_name_text, intern_font, font_size, intern_y_adjust, intern_x_adjust)

        date_font_size = 1.5
        date_font = cv2.FONT_HERSHEY_SIMPLEX
        # For from Date
        from_date_text = '25 Mar 22'
        # from_date_text = str(row["from"]) + " " + str(row["s_month"]) + " " + str(row["s_year"])
        print(f"from: {from_date_text}")
        f_date_x_adjust, date_y_adjust = 130, -80
        draw_text(img, from_date_text, date_font, date_font_size, f_date_x_adjust, date_y_adjust)

        # for to Date
        to_date_x_adjust = 535
        to_date_text = '14 Apr 22'
        # to_date_text = str(row["to"]) + " " + str(row["e_month"]) + " " + str(row["e_year"])
        print(f"to: {to_date_text}")
        draw_text(img, to_date_text, date_font, date_font_size, to_date_x_adjust, date_y_adjust)
        file_path = "Internship_Completion_Certificate " + row["Name"]

        certi_path = output_directory_path + file_path + ".png"
        print(certi_path)
        cv2.imwrite(certi_path, img)

        pdf_path = output_directory_path + file_path + ".pdf"
        print(pdf_path)
        with open(pdf_path, "wb") as f:
            f.write(img2pdf.convert(certi_path))


if __name__ == '__main__':
    main_function()
    cv2.destroyAllWindows()


