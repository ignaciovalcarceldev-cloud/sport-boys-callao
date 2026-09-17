import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

doc = Document()

style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)

for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3)
    section.right_margin = Cm(3)

def add_heading_styled(text, level=1):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 51, 102)
    return heading

def add_table_bordered(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    borders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        borders.append(border)
    tblPr.append(borders)

# PORTADA
doc.add_paragraph('')
doc.add_paragraph('')
doc.add_paragraph('')

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('1AAD2788 - Organización y Dirección de Empresas')
run.bold = True
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0, 51, 102)

doc.add_paragraph('')

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('Primer Avance del Proyecto (TB1)')
run.bold = True
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(0, 51, 102)

doc.add_paragraph('')

project = doc.add_paragraph()
project.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = project.add_run('NexoraAI - Asistente Inteligente de Ventas y Seguimiento\nde Asistencia para MYPES en el Perú')
run.font.size = Pt(13)
run.font.color.rgb = RGBColor(0, 102, 153)

doc.add_paragraph('')
doc.add_paragraph('')

info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = info.add_run(f'Alumno: Ignacio Valcarcel Llerena\nCiclo: 2026-10\nFecha: {datetime.date.today().strftime("%d/%m/%Y")}')
run.font.size = Pt(11)

doc.add_page_break()

# I. IDENTIFICACIÓN Y SELECCIÓN DE LA IDEA DE NEGOCIO
add_heading_styled('I. IDENTIFICACIÓN Y SELECCIÓN DE LA IDEA DE NEGOCIO', level=1)

add_heading_styled('1.1 Perfil del Emprendedor', level=2)
doc.add_paragraph(
    'El equipo emprendedor está conformado por estudiantes de Ingeniería de Sistemas de la Universidad Peruana de Ciencias Aplicadas (UPC), '
    'con formación en desarrollo de software, inteligencia artificial y gestión empresarial. El equipo combina conocimientos técnicos en programación, '
    'análisis de datos y arquitectura de software con una comprensión del entorno empresarial peruano, especialmente el sector MYPE.'
)
doc.add_paragraph(
    'Las principales fortalezas del equipo incluyen:'
)
items = [
    'Capacidad técnica en desarrollo de software y modelos de inteligencia artificial.',
    'Conocimiento del ecosistema MYPE peruano y sus desafíos digitales.',
    'Habilidades de investigación de mercado y validación de hipótesis.',
    'Compromiso con la creación de soluciones que generen impacto social positivo.',
    'Experiencia previa en proyectos de tecnología aplicada a problemas reales.'
]
for item in items:
    doc.add_paragraph(item, style='List Bullet')

add_heading_styled('1.2 Generación de Ideas', level=2)
doc.add_paragraph(
    'Para la generación de ideas se aplicaron técnicas de creatividad como Brainstorming y análisis de tendencias tecnológicas. '
    'Se identificaron múltiples oportunidades de negocio basadas en la implementación de inteligencia artificial en el mercado peruano. '
    'A continuación se presenta la matriz de ranking de las ideas evaluadas:'
)

table = doc.add_table(rows=6, cols=6)
add_table_bordered(table)
table.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['Idea', 'Relevancia\n(1-10)', 'Viabilidad\n(1-10)', 'Escalabilidad\n(1-10)', 'Impacto Social\n(1-10)', 'Puntaje\nTotal']
for i, header in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.text = header
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.bold = True
            run.font.size = Pt(9)

ideas_data = [
    ['Asistente IA de ventas y seguimiento para MYPES de servicios', '10', '9', '9', '9', '37'],
    ['Plataforma de diagnóstico empresarial con IA para MYPES', '8', '8', '9', '8', '33'],
    ['Chatbot de atención al cliente para comercio electrónico', '9', '7', '8', '6', '30'],
    ['Sistema de predicción de demanda para pequeños comercios', '7', '7', '8', '7', '29'],
    ['Plataforma de capacitación empresarial con IA adaptativa', '8', '6', '7', '8', '29'],
]
for row_idx, row_data in enumerate(ideas_data, 1):
    for col_idx, cell_data in enumerate(row_data):
        cell = table.rows[row_idx].cells[col_idx]
        cell.text = cell_data
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(9)

doc.add_paragraph('')
doc.add_paragraph(
    'La idea con mayor puntaje fue "Asistente IA de ventas y seguimiento para MYPES de servicios" con 37 puntos, '
    'destacando por su alta relevancia en el mercado actual, viabilidad técnica, potencial de escalabilidad y impacto social.'
)

add_heading_styled('1.3 Mejora de Ideas: Técnica SCAMPER', level=2)
doc.add_paragraph(
    'Se aplicó la técnica SCAMPER sobre la idea seleccionada para identificar mejoras y diferenciadores competitivos:'
)

scamper_data = [
    ['Sustituir', 'Sustituir el CRM tradicional por un asistente conversacional basado en IA que opere a través de WhatsApp, '
     'el canal que las MYPES ya utilizan diariamente.'],
    ['Combinar', 'Combinar funciones de gestión de clientes, seguimiento automatizado, generación de cotizaciones y análisis '
     'predictivo en una sola plataforma accesible desde el celular.'],
    ['Adaptar', 'Adaptar la solución al contexto peruano: integración con Yape/Plin para confirmaciones de pago, '
     'soporte en español con modismos locales, y compatibilidad con la facturación de la SUNAT.'],
    ['Modificar', 'Modificar el enfoque de los CRMs tradicionales: en lugar de forzar al usuario a aprender una interfaz compleja, '
     'el asistente interactúa mediante conversación natural por WhatsApp.'],
    ['Poner en otros usos', 'Aplicar la tecnología de modelos de lenguaje (LLM) para automatizar no solo la atención al cliente, '
     'sino también la generación de reportes de ventas y la predicción de oportunidades.'],
    ['Eliminar', 'Eliminar la necesidad de infraestructura tecnológica compleja, capacitación extensa y personal técnico dedicado. '
     'El asistente funciona desde un teléfono celular con conexión a internet.'],
    ['Reordenar', 'Reordenar el flujo comercial: en lugar de que el dueño recuerde hacer seguimiento, el sistema lo hace '
     'automáticamente y lo notifica cuando hay una oportunidad que requiere acción humana.'],
]
for element, description in scamper_data:
    p = doc.add_paragraph()
    run = p.add_run(f'{element}: ')
    run.bold = True
    p.add_run(description)

add_heading_styled('1.4 Justificación de la Idea Prioritaria', level=2)
doc.add_paragraph(
    'La selección de un asistente inteligente de ventas y seguimiento para MYPES se sustenta en datos concretos del mercado peruano:'
)

doc.add_paragraph('')
p = doc.add_paragraph()
run = p.add_run('Dimensiones del mercado MYPE:')
run.bold = True

doc.add_paragraph(
    'Según el Ministerio de la Producción (PRODUCE), en 2026 existen aproximadamente 2.44 millones de MYPE formales en el Perú, '
    'representando el 99.2% de las empresas formales del país. Estas generan alrededor de 10.3 millones de empleos y representan '
    'el 20.6% del PBI nacional (PRODUCE, 2026).'
)

p = doc.add_paragraph()
run = p.add_run('Brecha digital:')
run.bold = True

doc.add_paragraph(
    'PRODUCE informó en julio de 2026 que aproximadamente el 30% de las MYPE ha avanzado significativamente en procesos de '
    'transformación digital, mientras que cerca del 70% presenta un rezago importante en madurez digital. Al mismo tiempo, '
    '9 de cada 10 MYPE cuentan con acceso a Internet, pero solo alrededor del 7% realiza ventas por Internet '
    '(PRODUCE, "Las MYPE que se digitalizan tienen el potencial de incrementar sus ventas hasta en un 25%", 16 de julio de 2026).'
)

p = doc.add_paragraph()
run = p.add_run('Adopción de IA en Perú:')
run.bold = True

doc.add_paragraph(
    'Según un estudio de Experis Perú, 8 de cada 10 empresas peruanas integraron inteligencia artificial en sus operaciones '
    'al cierre de 2025. Las empresas peruanas proyectan un crecimiento de 3.9 veces en inversión en IA dentro de su gasto total '
    'en tecnología, el incremento más alto de Latinoamérica según Lenovo e IDC. El retorno promedio es de 2.5 veces la inversión '
    'inicial, con un tiempo de recuperación de 13.7 meses (Ecosistema Startup, julio 2026).'
)

p = doc.add_paragraph()
run = p.add_run('Acceso económico:')
run.bold = True

doc.add_paragraph(
    'Las herramientas de IA para MYPEs están disponibles desde S/60 mensuales, democratizando el acceso a tecnología que '
    'anteriormente era exclusiva de grandes corporaciones. El mercado de herramientas SaaS ha eliminado la necesidad de contratar '
    'desarrolladores o invertir miles de dólares (Ecosistema Startup, 2026).'
)

p = doc.add_paragraph()
run = p.add_run('Problema específico que resuelve:')
run.bold = True

doc.add_paragraph(
    'Las pequeñas empresas pierden oportunidades comerciales porque la información de sus clientes está dispersa entre WhatsApp, '
    'llamadas, hojas de cálculo, correo electrónico y otros canales. El seguimiento depende de la memoria del dueño o de trabajadores '
    'individuales. Un flujo típico incluye: un cliente pregunta por un servicio, la empresa responde, prepara una cotización, espera '
    'la respuesta y luego debe recordar hacer seguimiento. Sin un proceso organizado, las oportunidades quedan olvidadas.'
)

p = doc.add_paragraph()
run = p.add_run('Cifras de impacto:')
run.bold = True

doc.add_paragraph(
    'La herramienta MAIA (MYPE AsesorIA) de PRODUCE demuestra que existe interés institucional y empresarial en utilizar IA a través '
    'de canales que los pequeños negocios ya conocen como WhatsApp. Sin embargo, MAIA es una herramienta de orientación, no un sistema '
    'operativo comercial. La oportunidad privada está en conectar conversaciones, clientes, cotizaciones, seguimiento y tareas dentro '
    'de un flujo concreto de negocio.'
)

doc.add_paragraph(
    'El mercado de CRM en Perú está en crecimiento acelerado. Según 6Wresearch, el mercado de CRM en Perú proyecta crecimiento '
    'sostenido hasta 2032, impulsado por la adopción de soluciones cloud y la integración de IA. Plataformas como Kommo (USD 15/usuario/mes), '
    'Whaticket (USD 49/mes) y Simla.com (USD 79/mes) ya compiten en el mercado peruano, validando la demanda de soluciones de gestión '
    'de relaciones con clientes para pymes (Elige tu CRM, 2026).'
)

doc.add_paragraph(
    'La oportunidad más interesante en esta etapa es comenzar con un producto específico: un asistente de ventas y seguimiento para '
    'pequeñas empresas de servicios o B2B, que reciba consultas de clientes, organice la información del lead, ayude a preparar '
    'respuestas o cotizaciones, recuerde el seguimiento y muestre el estado de cada oportunidad, manteniendo la automatización bajo '
    'supervisión humana al inicio.'
)

doc.add_page_break()

# II. MODELO DE NEGOCIO
add_heading_styled('II. MODELO DE NEGOCIO', level=1)

add_heading_styled('2.1 Business Model Canvas (BMC)', level=2)

bmc_items = [
    ('Socios Clave', 
     '• Proveedores de APIs de IA (OpenAI, Google, Anthropic)\n'
     '• Proveedores de servicios de mensajería (WhatsApp Business API)\n'
     '• ProInnóvate y organismos de apoyo a MYPES\n'
     '• Aliados tecnológicos (AWS, Google Cloud)\n'
     '• Cámaras de comercio y gremiales empresariales'),
    ('Actividades Clave',
     '• Desarrollo y mantenimiento del asistente IA\n'
     '• Investigación de mercado y validación con usuarios\n'
     '• Atención al cliente y soporte técnico\n'
     '• Marketing digital y adquisición de clientes\n'
     '• Análisis de datos y mejora continua del producto'),
    ('Recursos Clave',
     '• Plataforma de IA basada en modelos de lenguaje\n'
     '• Infraestructura en la nube (AWS/GCP)\n'
     '• Equipo de desarrollo de software\n'
     '• Base de conocimiento del sector MYPE\n'
     '• Datos de entrenamiento y retroalimentación de usuarios'),
    ('Propuesta de Valor',
     'Asistente inteligente que centraliza la gestión comercial de MYPES a través de WhatsApp, '
     'automatizando el seguimiento de clientes, generación de cotizaciones y análisis de oportunidades '
     'de venta, sin requerir conocimientos técnicos ni infraestructura adicional.'),
    ('Relaciones con Clientes',
     '• Atención personalizada por WhatsApp y correo\n'
     '• Onboarding guiado paso a paso\n'
     '• Soporte técnico continuo\n'
     '• Comunidad de usuarios y mejores prácticas\n'
     '• Actualizaciones y mejoras basadas en feedback'),
    ('Canales',
     '• WhatsApp Business como canal principal\n'
     '• Sitio web institucional\n'
     '• Redes sociales (Instagram, Facebook, LinkedIn)\n'
     '• Aliados comerciales y gremiales\n'
     '• Ferias y eventos empresariales'),
    ('Segmentos de Clientes',
     '• MYPES de servicios técnicos (electricistas, plomeros, instaladores)\n'
     '• Distribuidores B2B\n'
     '• Talleres y servicios automotrices\n'
     '• Inmobiliarias y agencias inmobiliarias\n'
     '• Pequeños comercios y tiendas'),
    ('Estructura de Costos',
     '• Infraestructura cloud (servidores, APIs)\n'
     '• Desarrollo y mantenimiento de software\n'
     '• Costo de APIs de IA por uso\n'
     '• Marketing y adquisición de clientes\n'
     '• Equipo humano (desarrollo, soporte, ventas)'),
    ('Fuentes de Ingreso',
     '• Suscripción mensual por usuario (desde S/120/mes)\n'
     '• Planes por número de contactos gestionados\n'
     '• Comisiones por transacciones procesadas\n'
     '• Servicios de implementación y capacitación\n'
     '• Planes enterprise para empresas medianas'),
]

bmc_table = doc.add_table(rows=3, cols=3)
add_table_bordered(bmc_table)

# First row: 3 items
for i, (title, content) in enumerate(bmc_items[:3]):
    cell = bmc_table.rows[0].cells[i]
    p = cell.paragraphs[0]
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(9)
    p2 = cell.add_paragraph(content)
    p2.style.font.size = Pt(8)

# Second row: 2 items + empty
for i, (title, content) in enumerate(bmc_items[3:5]):
    cell = bmc_table.rows[1].cells[i]
    p = cell.paragraphs[0]
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(9)
    p2 = cell.add_paragraph(content)
    p2.style.font.size = Pt(8)

# Third row: propuesta de valor (center)
cell = bmc_table.rows[1].cells[2]
p = cell.paragraphs[0]
run = p.add_run('Propuesta de Valor')
run.bold = True
run.font.size = Pt(9)
cell.add_paragraph(bmc_items[3][1]).style.font.size = Pt(8)

# Row 3: 3 items
for i, (title, content) in enumerate(bmc_items[5:8]):
    cell = bmc_table.rows[2].cells[i]
    p = cell.paragraphs[0]
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(9)
    p2 = cell.add_paragraph(content)
    p2.style.font.size = Pt(8)

# Last row: fuentes de ingreso
cell = bmc_table.rows[2].cells[2]
p = cell.paragraphs[0]
run = p.add_run('Fuentes de Ingreso')
run.bold = True
run.font.size = Pt(9)
cell.add_paragraph(bmc_items[8][1]).style.font.size = Pt(8)

doc.add_paragraph('')

doc.add_page_break()

# III. PLANEAMIENTO ESTRATÉGICO
add_heading_styled('III. PLANEAMIENTO ESTRATÉGICO', level=1)

add_heading_styled('3.1 Misión', level=2)
doc.add_paragraph(
    'Nuestra misión es empoderar a las micro, pequeñas y medianas empresas del Perú mediante soluciones de inteligencia artificial '
    'accesibles y prácticas que les permitan organizar su gestión comercial, convertir consultas en ventas y crecer de manera sostenible, '
    'contribuyendo al desarrollo económico y social de sus comunidades.'
)

add_heading_styled('3.2 Visión', level=2)
doc.add_paragraph(
    'Ser para 2030 la plataforma de asistencia comercial con inteligencia artificial de referencia para las MYPES en Perú y Latinoamérica, '
    'operando en al menos cinco países, con cien mil agentes activos y una comunidad de diez mil profesionales que utilicen nuestra '
    'tecnología para impulsar la competitividad de sus negocios.'
)

add_heading_styled('3.3 Valores', level=2)
valores = [
    ('Accesibilidad', 'Creemos que la tecnología de inteligencia artificial debe estar al alcance de todos los negocios, sin importar su tamaño o recursos tecnológicos.'),
    ('Simplicidad', 'Diseñamos soluciones intuitivas que eliminan la complejidad técnica y permiten al usuario enfocarse en lo que más importa: vender y hacer crecer su negocio.'),
    ('Confiabilidad', 'Garantizamos la seguridad y privacidad de los datos de nuestros clientes, manteniendo estándares éticos en el uso de la inteligencia artificial.'),
    ('Impacto Social', 'Nuestra tecnología busca no solo generar rentabilidad, sino también contribuir al fortalecimiento del tejido empresarial peruano.'),
    ('Innovación Continua', 'Nos comprometemos a mejorar constantemente nuestra solución basándonos en las necesidades reales de los usuarios y las tendencias tecnológicas.'),
]
for valor, desc in valores:
    p = doc.add_paragraph()
    run = p.add_run(f'{valor}: ')
    run.bold = True
    p.add_run(desc)

add_heading_styled('3.4 Objetivos Estratégicos (SMART)', level=2)

smart_data = [
    ['Objetivo', 'Específico', 'Medible', 'Alcanzable', 'Relevante', 'Temporal'],
    ['Validar la propuesta de valor',
     'Realizar entrevistas con MYPES de servicios en Lima para validar el problema y la solución propuesta',
     'Minimum 50 entrevistas realizadas con empresarios del sector',
     'Equipo con capacidad de contacto directo y acceso a gremiales',
     'Fundamenta la decisión de construir o pivotar el producto',
     'En los primeros 3 meses'],
    ['Desarrollar un MVP funcional',
     'Construir un asistente conversacional por WhatsApp que registre leads, genere cotizaciones básicas y programe seguimientos',
     'MVP funcional con al menos 3 features principales operativas',
     'Equipo técnico con experiencia en IA y desarrollo de software',
     'Permite probar la solución con usuarios reales',
     'En 4 meses desde el inicio'],
    ['Adquirir los primeros 20 clientes piloto',
     'Captar 20 MYPES de servicios que utilicen el asistente durante 3 meses como piloto',
     '20 clientes activos con uso mínimo semanal del sistema',
     'Red de contactos y alianzas con gremiales locales',
     'Genera ingresos iniciales y retroalimentación para mejorar el producto',
     'En 6 meses'],
    ['Lograr una tasa de retención del 70%',
     'Mantener al menos 14 de los 20 clientes piloto activos después de 3 meses de uso',
     'Retención ≥ 70% medido por uso mensual activo',
     'Mejora continua del producto basada en feedback',
     'Demuestra que la solución genera valor real para el usuario',
     'En 9 meses'],
    ['Obtener inversión seed de hasta S/150,000',
     'Presentar pitch a inversores angel y aceleradoras con métricas de tracción del piloto',
     'Mínimo 10 reuniones con inversores y 1 term sheet recibido',
     'Tracción demostrada y equipo comprometido',
     'Permite escalar el producto y la operación comercial',
     'En 12 meses'],
]

smart_table = doc.add_table(rows=len(smart_data), cols=6)
add_table_bordered(smart_table)
for row_idx, row_data in enumerate(smart_data):
    for col_idx, cell_data in enumerate(row_data):
        cell = smart_table.rows[row_idx].cells[col_idx]
        cell.text = cell_data
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(8)
                if row_idx == 0:
                    run.bold = True

doc.add_paragraph('')

add_heading_styled('3.5 Análisis Externo', level=2)

add_heading_styled('3.5.1 Análisis PESTEL', level=3)

pestel_data = [
    ['Factor', 'Análisis'],
    ['Político',
     'El Gobierno del Perú aprobó el Reglamento de la Ley N° 31814 (septiembre 2025) para el uso de IA. '
     'La Estrategia Nacional de Inteligencia Artificial 2026-2030 (ENIA) establece metas de despliegue, '
     'capacitación y sandboxes regulatorios. ProInnóvate financia startups de IA para MYPES.'],
    ['Económico',
     'Las MYPE representan el 99.2% de empresas formales y el 20.6% del PBI. Las empresas peruanas proyectan '
     'un crecimiento de 3.9x en inversión en IA. El retorno promedio de IA es de 2.5x en 13.7 meses. '
     'Las herramientas de IA están disponibles desde S/60/mes.'],
    ['Social',
     '87% de pymes está lista para adoptar IA (Informe ASUS 2025). WhatsApp es el canal dominante con 99% de '
     'penetración. La cultura empresarial peruana está evolucionando hacia la adopción tecnológica, '
     'especialmente en MYPES jóvenes y urbanas.'],
    ['Tecnológico',
     '8 de cada 10 empresas integraron IA al cierre de 2025. El 51% implementó IA en menos de 6 meses. '
     'Existe infraestructura cloud accesible (AWS, GCP). Los modelos de lenguaje (LLM) permiten crear '
     'asistentes conversacionales avanzados a bajo costo.'],
    ['Ecológico',
     'La solución opera en la nube, minimizando la huella de carbono. Se promueve el uso eficiente de '
     'recursos tecnológicos y la digitalización como alternativa a procesos presenciales.'],
    ['Legal',
     'Ley N° 31814 regula el uso de IA en Perú. La ENIA establece marcos de gobernanza algorítmica. '
     'Se requiere garantizar supervisión humana, transparencia en decisiones automatizadas y protección '
     'de datos personales conforme a la normativa vigente.'],
]

pestel_table = doc.add_table(rows=len(pestel_data), cols=2)
add_table_bordered(pestel_table)
for row_idx, row_data in enumerate(pestel_data):
    for col_idx, cell_data in enumerate(row_data):
        cell = pestel_table.rows[row_idx].cells[col_idx]
        cell.text = cell_data
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)
                if row_idx == 0:
                    run.bold = True

doc.add_paragraph('')

add_heading_styled('3.5.2 Análisis de las 5 Fuerzas de Porter', level=3)

porter_data = [
    ['Fuerza', 'Nivel', 'Análisis'],
    ['Poder de negociación de proveedores',
     'Medio',
     'Los proveedores de APIs de IA (OpenAI, Google) tienen cierto poder por ser pocos, pero la competencia '
     'creciente reduce su capacidad de fijación de precios. Los proveedores de cloud (AWS, GCP) ofrecen '
     'planes accesibles para startups.'],
    ['Poder de negociación de compradores',
     'Alto',
     'Las MYPES tienen múltiples opciones de CRM y herramientas digitales. Sin embargo, la mayoría son '
     'complejas o costosas. Un producto simple y accesible puede capturar este segmento.'],
    ['Amenaza de nuevos entrantes',
     'Alto',
     'El mercado de herramientas IA para pymes está en crecimiento. Existen competidores como MaravIA, '
     'Yala y Metric 360. La barrera de entrada técnica es moderada, pero la diferenciación por '
     'conocimiento del mercado local crea ventaja.'],
    ['Amenaza de productos sustitutos',
     'Medio',
     'Los CRMs tradicionales (HubSpot, Zoho, Pipedrive) son sustitutos potenciales, pero suelen ser '
     'complejos y poco intuitivos para MYPES. Las hojas de cálculo y la gestión manual también son '
     'sustitutos, pero ineficientes.'],
    ['Rivalidad competitiva',
     'Media-Alta',
     'El mercado peruano tiene competidores locales (MaravIA, Yala) e internacionales (Kommo, Whaticket). '
     'Sin embargo, la mayoría se enfoca en generalidades. Un nicho específico de servicios puede '
     'reducir la rivalidad directa.'],
]

porter_table = doc.add_table(rows=len(porter_data), cols=3)
add_table_bordered(porter_table)
for row_idx, row_data in enumerate(porter_data):
    for col_idx, cell_data in enumerate(row_data):
        cell = porter_table.rows[row_idx].cells[col_idx]
        cell.text = cell_data
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)
                if row_idx == 0:
                    run.bold = True

doc.add_paragraph('')

add_heading_styled('3.6 Análisis Interno', level=2)

doc.add_paragraph(
    'El análisis interno de la startup NexoraAI considera los siguientes elementos:'
)

foda_interno = [
    ('Recursos Humanos',
     'Equipo multidisciplinario con formación en ingeniería de sistemas, conocimiento de IA y '
     'compromiso con la solución de problemas reales del mercado peruano.'),
    ('Tecnología',
     'Experiencia en desarrollo de software, modelos de lenguaje (LLM), integración con APIs '
     'de mensajería y arquitectura de nube. Capacidad de prototipado rápido.'),
    ('Conocimiento del Mercado',
     'Investigación previa del sector MYPE peruano, comprensión de las brechas digitales '
     'y experiencia en contacto directo con empresarios del sector.'),
    ('Capacidad Financiera',
     'Startup en etapa temprana con acceso potencial a fondos de ProInnóvate y programase de '
     'aceleración. Necesidad de inversión seed para escalar.'),
    ('Marca y Posicionamiento',
     'Startup en construcción. Sin presencia de mercado aún. Oportunidad de posicionarse '
     'como referente en IA aplicada a MYPES de servicios en Perú.'),
]

for recurso, desc in foda_interno:
    p = doc.add_paragraph()
    run = p.add_run(f'{recurso}: ')
    run.bold = True
    p.add_run(desc)

add_heading_styled('3.7 Matriz FODA', level=2)

foda_data = [
    ['', 'Factores Positivos', 'Factores Negativos'],
    ['Factores\nInternos',
     'Fortalezas:\n'
     '• Equipo técnico con conocimiento en IA\n'
     '• Investigación de mercado previa\n'
     '• Compromiso con impacto social\n'
     '• Solución innovadora y diferenciada\n'
     '• Bajos costos operativos iniciales',
     'Debilidades:\n'
     '• Startup en etapa temprana\n'
     '• Sin presencia de mercado\n'
     '• Recursos financieros limitados\n'
     '• Dependencia de proveedores de IA\n'
     '• Equipo pequeño'],
    ['Factores\nExternos',
     'Oportunidades:\n'
     '• Mercado MYPE de 2.44M de empresas\n'
     '• 70% de MYPEs sin transformación digital\n'
     '• Ley 31814 y ENIA 2026-2030\n'
     '• Fondos de ProInnóvate disponibles\n'
     '• Retorno promedio de IA: 2.5x\n'
     '• Herramientas desde S/60/mes',
     'Amenazas:\n'
     '• Competencia de CRMs internacionales\n'
     '• Barreras de adopción tecnológica\n'
     '• Regulación de IA en evolución\n'
     '• Riesgo de imitación por grandes actores\n'
     '• Dependencia de WhatsApp Business API'],
]

foda_table = doc.add_table(rows=3, cols=3)
add_table_bordered(foda_table)
for row_idx, row_data in enumerate(foda_data):
    for col_idx, cell_data in enumerate(row_data):
        cell = foda_table.rows[row_idx].cells[col_idx]
        cell.text = cell_data
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)
                if row_idx == 0 or col_idx == 0:
                    run.bold = True

doc.add_page_break()

# IV. MARCO ÉTICO-LEGAL Y RESPONSABILIDAD SOCIAL
add_heading_styled('IV. MARCO ÉTICO-LEGAL Y RESPONSABILIDAD SOCIAL (RSE)', level=1)

add_heading_styled('4.1 Análisis de Sensibilidad Social', level=2)

doc.add_paragraph(
    'El negocio NexoraAI responde a necesidades reales de la sociedad peruana en múltiples dimensiones:'
)

p = doc.add_paragraph()
run = p.add_run('Democratización del acceso tecnológico: ')
run.bold = True
p.add_run(
    'Las MYPES representan el 99.2% de las empresas formales del Perú y generan más de 10 millones de empleos, '
    'sin embargo, el 70% presenta rezago en transformación digital. NexoraAI busca cerrar esta brecha '
    'ofreciendo una herramienta de IA accesible, desde S/120 mensuales, sin requerir conocimientos técnicos.'
)

p = doc.add_paragraph()
run = p.add_run('Generación de empleo productivo: ')
run.bold = True
p.add_run(
    'Al mejorar la gestión comercial de las MYPES, el asistente inteligente contribuye a que estos negocios '
    'crezcan, generen más ventas y contraten más personal. Cada MYPE que mejora su productividad potencialmente '
    'crea empleos directos e indirectos en su comunidad.'
)

p = doc.add_paragraph()
run = p.add_run('Reducción de la desigualdad económica: ')
run.bold = True
p.add_run(
    'Las pequeñas empresas en zonas alejadas de Lima enfrentan mayores barreras para acceder a herramientas '
    'tecnológicas. Al operar a través de WhatsApp, NexoraAI llega a cualquier rincón del país donde exista '
    'conexión a Internet, contribuyendo a reducir la brecha entre Lima y las regiones.'
)

p = doc.add_paragraph()
run = p.add_run('Fortalecimiento del tejido empresarial local: ')
run.bold = True
p.add_run(
    'El 80% de las empresas peruanas mueren en su primer año por desorden administrativo (MaravIA, 2025). '
    'NexoraAI contribuye a la supervivencia de las MYPES proporcionando herramientas que anteriormente '
    'estaban reservadas para grandes empresas, fortaleciendo así la economía local.'
)

p = doc.add_paragraph()
run = p.add_run('Inclusión financiera: ')
run.bold = True
p.add_run(
    'Al integrar confirmaciones de pago vía Yape/Plin y gestionar cotizaciones de forma digital, '
    'NexoraAI acerca a las MYPES al ecosistema financiero digital, facilitando transparencia y trazabilidad '
    'en sus operaciones comerciales.'
)

add_heading_styled('4.2 Integración Ética', level=2)

doc.add_paragraph(
    'La toma de decisiones en NexoraAI incorpora el enfoque solidario y el respeto a los deberes y '
    'derechos ciudadanos de la siguiente manera:'
)

p = doc.add_paragraph()
run = p.add_run('Transparencia algorítmica: ')
run.bold = True
p.add_run(
    'El asistente inteligente opera bajo principios de transparencia. Cada recomendación, cotización o '
    'seguimiento generado por la IA es explicitable y puede ser revisado por el usuario humano antes de '
    'ser enviado al cliente final. Nunca se toman decisiones comerciales sin supervisión humana.'
)

p = doc.add_paragraph()
run = p.add_run('Protección de datos personales: ')
run.bold = True
p.add_run(
    'Todos los datos de clientes y conversaciones se almacenan con cifrado de extremo a extremo. '
    'Se cumple con la Ley de Protección de Datos Personales del Perú (Ley N° 29733) y se garantiza '
    'que la información no sea utilizada para fines diferentes a los contratados por el usuario.'
)

p = doc.add_paragraph()
run = p.add_run('Uso responsable de la IA: ')
run.bold = True
p.add_run(
    'Siguiendo el Reglamento de la Ley N° 31814, NexoraAI incorpora supervisión humana en todas las '
    'acciones automatizadas, documenta las decisiones del sistema y establece mecanismos de auditoría. '
    'El usuario siempre tiene la última palabra en las interacciones comerciales.'
)

p = doc.add_paragraph()
run = p.add_run('No discriminación: ')
run.bold = True
p.add_run(
    'El asistente está diseñado para ser accesible a usuarios de todos los niveles educativos y '
    'tecnológicos. La interfaz conversacional por WhatsApp elimina barreras de alfabetización digital '
    'y no discrimina por ubicación geográfica, género o nivel socioeconómico.'
)

p = doc.add_paragraph()
run = p.add_run('Compromiso con el bien común: ')
run.bold = True
p.add_run(
    'NexoraAI se compromete a destinar un porcentaje de sus ingresos a programas de capacitación '
    'digital gratuita para MYPES en zonas rurales y marginadas, contribuyendo al desarrollo '
    'sostenible del país.'
)

p = doc.add_paragraph()
run = p.add_run('Impacto ambiental: ')
run.bold = True
p.add_run(
    'Al operar en la nube y promover la digitalización de procesos, NexoraAI reduce la necesidad '
    'de desplazamientos presenciales, el uso de papel y la generación de residuos. El asistente '
    'contribuye a un modelo de negocio más sostenible y consciente con el medio ambiente.'
)

doc.add_page_break()

# CONCLUSIONES
add_heading_styled('CONCLUSIONES', level=1)

doc.add_paragraph(
    'El mercado peruano presenta una oportunidad significativa para soluciones de software orientadas a MYPES '
    'debido a su enorme cantidad de empresas (2.44 millones de MYPES formales) y a la brecha existente entre '
    'acceso a Internet (90%) y transformación digital (solo 30% avanzado). PRODUCE está impulsando activamente '
    'comercio electrónico, automatización, análisis de datos e inteligencia artificial, lo que refuerza la '
    'relevancia del problema identificado.'
)

doc.add_paragraph(
    'La mejor hipótesis actual es una solución de IA enfocada inicialmente en ventas y seguimiento para un '
    'nicho específico de empresas de servicios o B2B. Los datos del mercado respaldan esta decisión: '
    'el 80% de las empresas ya usa IA, el retorno promedio es de 2.5x, y las herramientas están disponibles '
    'desde S/60/mes. Sin embargo, esta conclusión no debe considerarse definitiva.'
)

doc.add_paragraph(
    'El siguiente paso es investigación primaria: hablar con empresarios, observar sus procesos y medir '
    'cuánto les cuesta realmente el problema de la información dispersa y la falta de seguimiento. '
    'Si los datos confirman la hipótesis, se puede diseñar un MVP en 4 meses. Si no la confirman, '
    'se cambia de problema antes de invertir tiempo considerable en desarrollo.'
)

# FUENTES
doc.add_page_break()
add_heading_styled('REFERENCIAS BIBLIOGRÁFICAS', level=1)

fuentes = [
    'Ministerio de la Producción (PRODUCE). "Las MYPE que se digitalizan tienen el potencial de incrementar sus ventas hasta en un 25%". 16 de julio de 2026.',
    'Ministerio de la Producción (PRODUCE). "Digitaliza tu negocio y haz crecer tu MYPE - MYPE Digital". Actualización de junio de 2026.',
    'Ministerio de la Producción (PRODUCE). "PRODUCE impulsa una nueva etapa de MAIA para acercar sus servicios a más productores y emprendedores". 22 de junio de 2026.',
    'Ministerio de la Producción (PRODUCE). "Más de 100 MYPE fueron capacitadas en el uso de herramientas tecnológicas para incrementar sus ventas". 22 de mayo de 2026.',
    'Ministerio de la Producción (PRODUCE). "Cooperativas y MYPE reciben capacitación en IA, WhatsApp Business y TikTok para dar el salto digital". 23 de abril de 2026.',
    'Ministerio de la Producción (PRODUCE). "CyberWow: MYPE impulsadas por PRODUCE generaron ventas por más de S/ 141 mil". 4 de mayo de 2026.',
    'Ecosistema Startup. "IA para MYPES en Perú: 8 de 10 empresas ya la usan en 2026". Julio 2026.',
    'Gestión. "Empresas peruanas lideran crecimiento de inversión en IA". Mayo 2026.',
    'MaravIA. "MaravIA: Revolución IA en Perú - ERP impulsado por IA para MYPES". Noviembre 2025.',
    'Elige tu CRM. "Mejor CRM con WhatsApp en Perú 2026: Top 8 (precios en PEN)". Julio 2026.',
    'Start-Up.pe. "CRM para pequeñas empresas en Perú: comparativa de precios y beneficios". Marzo 2026.',
    '6Wresearch. "Peru CRM Market Strategic Insights & Opportunities 2026". Septiembre 2026.',
    'Microsoft. "Estudio de adopción de IA en empresas peruanas". 2025.',
    'Lenovo-IDC. "CIO Playbook 2025 - Perú". 2025.',
    'Experis Perú. "Estudio de integración de IA en empresas peruanas". 2025.',
]

for i, fuente in enumerate(fuentes, 1):
    doc.add_paragraph(f'{i}. {fuente}')

# GUARDAR
output_path = r'C:/Users/ignac/Downloads/TB1_NexoraAI_Completo.docx'
doc.save(output_path)
print(f'Documento guardado en: {output_path}')
