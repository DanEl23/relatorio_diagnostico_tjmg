# report_builder.py

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

def render(template_path, context, output_file):
    """
    Renderiza o template docx com o contexto e as imagens.
    """
    
    tpl = DocxTemplate(template_path)
    
    # Mapear imagens para inserção com tamanho (ex: 150mm de largura)
    img_context = {}
    if 'images' in context:
        for key, path in context['images'].items():
            img_context[key] = InlineImage(tpl, path, width=Mm(150))
    
    # Atualiza o contexto principal com as imagens preparadas
    context.update(img_context)
    
    print(f"Renderizando template: {template_path}...")
    
    try:
        tpl.render(context)
        tpl.save(output_file)
        print(f"Arquivo renderizado salvo em: {output_file}")
        return output_file
    except Exception as e:
        print(f"Erro ao renderizar o template: {e}")
        return None