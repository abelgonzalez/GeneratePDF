# Original source: https://medium.com/@theprasadpatil/how-to-create-a-pdf-report-from-excel-using-python-b882c725fcf6

# Import Dependencies
import os
import pandas as pd
import numpy as np
from fpdf import FPDF

# Reading input files
path = os.path.dirname(__file__)
Input_path = path+"/src/input/dataStream.xlsx"
df = pd.read_excel(Input_path, sheet_name='Sheet1')

# Creating variables

id = df['ID']
hr = df['HR']
rrInterval = df['RRInterval']
temperature = df['Temperature']
accelerometerX = df['AccelerometerX']
accelerometerY = df['AccelerometerY']
accelerometerZ = df['AccelerometerZ']
gyroX = df['GyroX']
gyroY = df['GyroY']
gyroZ = df['GyroZ']

#nickname = df['Nickname'].iloc[0]
#history = df['History'].iloc[0]
#cabinet = df['Cabinet'].iloc[0]
#ground = df['Stadium'].iloc[0]
#motto = df['Club Motto'].iloc[0]
#establishment = str(df['Founded in'].iloc[0])
#gaffer = df['Current Manager'].iloc[0]

# PDF creation
pdf = FPDF()  # Initializing object
pdf.add_page()  # Adding new page

# pdf.image(path + '/Football_field.svg.png', x=0,
#          y=0, w=210, h=300)  # Background wallpaper
# pdf.image(path + '/premier-league_logo.jpg', x=188,
#          y=0, w=18)  # Top Right corner logo

# Page Lineborders
pdf.set_line_width(0.0)
pdf.line(5.0,5.0,205.0,5.0) # top one
pdf.line(5.0,292.0,205.0,292.0) # bottom one
pdf.line(5.0,5.0,5.0,292.0) # left one
pdf.line(205.0,5.0,205.0,292.0) # right one

# Title of Document
#pdf.set_fill_color(200, 0, 0)
pdf.set_text_color(0, 0, 0)
pdf.set_font("Arial", size=20, style='B')
pdf.set_xy(10, 20)
pdf.cell(0, 12.5, txt='Monitoramento parâmetros fisiológicos',
         ln=1, align="C", fill=False)

# Text box of general information
pdf.set_fill_color(220,220,220) # grey color
#pdf.set_text_color(255, 255, 255)
pdf.set_font("Arial", size=15, style='B')
pdf.set_xy(10, 30)
pdf.cell(190, 10, txt='Informação geral', ln=1, align="L", fill=True)
# Text box of general information items
pdf.set_font("Arial", size=10, style='B')
pdf.cell(190, 5, txt='Data e horário de avaliação:', ln=1, align="L", fill=True)
pdf.cell(190, 5, txt='Usuário avaliado:', ln=1, align="L", fill=True)
pdf.cell(190, 5, txt='Idade:', ln=1, align="L", fill=True)


# Text box of Cardiac Frequency
#pdf.set_fill_color(220,220,220) # grey color
#pdf.set_text_color(255, 255, 255)
pdf.set_font("Arial", size=18, style='B')
pdf.set_xy(10, 70)
pdf.cell(190, 10, txt='FREQUÊNCIA CARDÍACA', ln=1, align="C", fill=False)
pdf.set_font("Arial", size=10, style='B') 

pdf.set_xy(10, 80)
pdf.cell(190, 5, txt='Método obtenção FC máxima:', ln=1, align="L", fill=False)
pdf.cell(190, 5, txt='Método uso de zonas de intensidade:', ln=1, align="L", fill=False)
pdf.set_xy(10, 90)
pdf.cell(190, 5, txt='FC_min:', ln=1, align="L", fill=False)
pdf.cell(190, 5, txt='FC_max:', ln=1, align="L", fill=False)
pdf.cell(190, 5, txt='Momentos acima do limiar 70%:', ln=1, align="L", fill=False)
pdf.cell(190, 5, txt='Momentos acima do limiar 85%:', ln=1, align="L", fill=False)
pdf.cell(190, 5, txt='Quanto tempo fica acima do limiar:', ln=1, align="L", fill=False)
pdf.cell(190, 5, txt='Quanto valor acima:', ln=1, align="L", fill=False)

# First and second Img
pdf.image(path + '/barchart.png', x=50, y=120, w=80)

#pdf.image(path + '/barchart.png', x=50, y=180, w=80)

# Temperature
pdf.set_font("Arial", size=18, style='B')
pdf.set_xy(10, 180)
pdf.cell(180, 10, txt='TEMPERATURA', ln=1, align="C", fill=False)
pdf.set_font("Arial", size=10, style='B') 

pdf.image(path + '/barchart.png', x=50, y=190, w=80)










# Text box of brief history
#pdf.set_text_color(0, 0, 0)
#pdf.set_font("Arial", size=9, style="B")
#pdf.set_xy(12, 76)
#pdf.multi_cell(120, 5, txt="brief", ln=1, align="J")

# Text box of header
#pdf.set_fill_color(0, 0, 150)
#pdf.set_text_color(255, 255, 255)
#pdf.set_font("Arial", size=20, style='B')
#pdf.set_xy(12, 153)
#pdf.cell(65, 10, txt='TROPHY CABINET', ln=1, align="L", fill=True)

# Text box of trophies won
#pdf.set_text_color(0, 0, 0)
#pdf.set_font("Arial", size=9, style="B")
#pdf.set_xy(12, 165)
#pdf.multi_cell(180, 5, txt="cabinet", ln=1, align="J")

# Club's Crest Image
#pdf.image(path + '/arsenal_logo.jpg', x=124, y=58, w=80)

# Club's Latin Motto
#pdf.set_text_color(220, 220, 220)
#pdf.set_font("Arial", size=14, style='B')
#pdf.set_xy(130, 130)
#pdf.cell(130, 10, txt="motto", ln=1, align="L")

# Inserting Text
#pdf.set_text_color(0, 0, 0)
#pdf.set_font("Arial", size=10.9)
#pdf.set_xy(145, 135)
#pdf.cell(130, 10, txt="Coach:", ln=1, align="L")

# Adding Manager's Name

#pdf.set_text_color(0, 0, 150)
#pdf.set_font("Arial", size=10.9, style='B')
#pdf.set_xy(158, 135)
#pdf.cell(130, 10, txt="gaffer", ln=1, align="L")


# Inserting stadium Picture
#pdf.image(path + '/Emirates.jpg', x=10, y=184, w=190)

# Stadium Name Text Box
#pdf.set_fill_color(255, 0, 0)
#pdf.set_text_color(255, 255, 250)
#pdf.set_font("Arial", size=15, style='B')
#pdf.set_xy(150, 260)
#pdf.cell(47, 8, txt="ground", ln=1, align="L", fill=True)

# Exporting PDF
pdf.output(path + "/report.pdf")

