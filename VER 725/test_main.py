# main.py
from html_to_docx_converter import HTMLToDocxConverter
from placeholder_replacer import apply_replacements
import json

def generate_document(html_template, replacement_data, output_filename):
    # Convert HTML to DOCX
    converter = HTMLToDocxConverter()
    doc = converter.convert(html_template)
    
    # Replace placeholders with data
    doc = apply_replacements(doc, replacement_data)
    
    # Save the final document
    doc.save(output_filename)
    print(f"Document saved as {output_filename}")

if __name__ == "__main__":
    # Load HTML template
    html_template = """
<p>  </p><table class="e-rte-table no-border" style="width: 100%; min-width: 0px;"><tbody><tr><td style="width: 50%;" class=""><p><span>РЕПУБЛИКА СРБИЈА</span></p><p><span style="background-color: yellow;">{{EXECUTOR_TITLE}}</span>&nbsp;<span style="background-color: yellow;">{{EXECUTOR_NAME}}</span>&nbsp;<br></p><p><span style="background-color: yellow;">{{JURISDICTION}}</span> </p><p><span style="background-color: yellow;">{{ADDRESS}}</span> <br></p><p>Телефон: <span style="background-color: yellow;">{{PHONE}}</span>&nbsp;<br></p><p>Посл. Бр. <span style="background-color: yellow;">{{CASE_NUMBER}}</span>&nbsp;<br></p><p>Дана&nbsp;<span style="background-color: yellow;">{{DATE}}</span>&nbsp; године</p><br></td><td style="width: 50%;" class=""><br></td></tr></tbody></table><p style="text-align: justify;"><span><span>Јавни извршитељ Филип Љујић, у извршном поступку извршног повериоца RAIFFEISEN BANKA AD BEOGRAD Z, Београд, ул. Ђорђа Станојевића бр. 16, МБ 17335600, ПИБ 100000299, број рачуна 908-0000000026501-15 који се води код банке НАРОДНА БАНКА СРБИЈЕ, против извршног дужника: Горан М. Бркић, Смедерево, ул. Милоја Ђака бр. 2/12, ЈМБГ 0209981370607, ради спровођења извршења одређеног Решењем о извршењу Основног суда у Смедереву 1ИВ-802/2013 од 27.08.2013.године, у поступку за намирење потраживања, доноси следећи:</span></span></p><h1 style="text-align: center;"><span style="font-size: 24pt;"><span>З А К Љ У Ч А К</span></span></h1><p style="text-align: center;"><span style="font-size: 14pt;">о спровођењу извршења на плати</span></p><p style="text-align: justify;"><span>I На основу Решења о извршењу Основног суда у Смедереву 1ИВ-802/2013 од 27.08.2013.године, ради намирења новчаног потраживања извршног повериоца и то:</span></p><p style="text-align: justify;"><br></p><ul><li style="text-align: justify;">у износу од <span style="text-decoration: underline;">68.255,01</span> динара на име главног дуга, са каматом обрачунатом у складу са Законом о затезној камати почев од 12.07.2013. године па до коначне исплате,</li><li style="text-align: justify;">у износу од <span style="text-decoration: underline;">39.056,00</span> динара на име трошкова извршења насталих пред судом,</li></ul><p style="text-align: justify;"><br></p><p style="text-align: justify;"><strong>ОДРЕЂУЈЕ СЕ ПЛЕНИДБА 2/3 ЗАРАДЕ</strong>, односно 1/2 месечне зараде, ако остварује минималну зараду, извршног дужника Горан М. Бркић, Смедерево, ул. Милоја Ђака бр. 2/12, ЈМБГ 0209981370607, коју извршном дужнику исплаћује послодавац "ГОРАН БРКИЋ ПР АГЕНЦИЈА ЗА КЊИГОВОДСТВЕНЕ УСЛУГЕ ПРОФИТ ГЕС" из Смедерева, ул.Балканска бр.100, МБ:64669478, ПИБ:110126987 и ИСПЛАТОМ на наменски рачун јавног извршитеља Филип Љујић ПР Јавни Извршитељ Смедерево, МБ 64331540, ПИБ 109650745 – број: <span style="text-decoration: underline;">205-0000000272303-16</span> код КОМЕРЦИЈАЛНА БАНКА А.Д. БЕОГРАД, са позивом на број предмета ИИВ 26/20.</p><p style="text-align: justify;"><br></p><p style="text-align: justify;"><strong>II НАЛАЖЕ СЕ </strong>послодавцу извршног дужника: "ГОРАН БРКИЋ ПР АГЕНЦИЈА ЗА КЊИГОВОДСТВЕНЕ УСЛУГЕ ПРОФИТ ГЕС" из Смедерева, ул.Балканска бр.100, МБ:64669478, ПИБ:110126987 , да износе из става I овог закључка ИСПЛАЋУЈЕ на наменски рачун јавног извршитеља и то редом, почев од трошкова поступка, износа камате и износа главног дуга, као и трошкова и награде јавног извршитеља, све до потпуног намирења потраживања извршног повериоца и јавног извршитеља по овом закључку.</p><p style="text-align: center;"><span style="font-size: 14pt;"><br></span></p>    """
    
    # Load replacement data
    replacement_data = json.loads('''
    {
    "COUNTRY": "РЕПУБЛИКА СРБИЈА",
    "EXECUTOR_NAME": "ФИЛИП ЉУЈИЋ",
    "EXECUTOR_TITLE": "ЈАВНИ ИЗВРШИТЕЉ",
    "JURISDICTION": "Именован за подручје Вишег суда у Смедереву и Привредног суда у Пожаревцу",
    "ADDRESS": "Смедерево, Карађорђева 32/2",
    "PHONE": "026/4103800",
    "CASE_NUMBER": "ИИВ 26/20",
    "DATE": "29.11.2020"
    }
    ''')
    
    # Generate the document
    generate_document(html_template, replacement_data, 'output_document.docx')