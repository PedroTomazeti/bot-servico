import re

def extrair_os_melhorada(descricao):
    """
    Extrai números de Ordem de Serviço (OS) de uma string de descrição.
    
    A função é otimizada para encontrar os seguintes formatos:
    1. OS individuais: "1234-ABC", "99-XYZ1"
    2. Múltiplas OS com sufixo comum: "123/124/125-DEF" ou "123_124-DEF"
    
    Retorna uma lista de strings com as OS encontradas, sem duplicatas.
    """
    # Usamos um set para garantir que os resultados sejam únicos e a inserção seja eficiente.
    os_encontradas = set()

    # --- ETAPA 1: Capturar o caso especial de múltiplos números com um único sufixo ---
    # Padrão para "0018/0019/20-COP" ou "123_456-ABC".
    # \b -> word boundary (fronteira de palavra), para não capturar parte de outro número.
    # ((?:\d{2,4}[/_])+\d{2,4}) -> Captura o grupo de números separados por / ou _.
    # -([A-Z0-9]{2,4}) -> Captura o sufixo alfanumérico.
    padrao_multiplo = re.compile(r'\b((?:\d{2,4}[/_])+\d{2,4})-([A-Z0-9]{2,4})\b')
    
    # Usamos finditer para ser mais eficiente em memória com strings grandes
    for match in padrao_multiplo.finditer(descricao):
        numeros_str, sufixo = match.groups()
        
        # Divide a string de números usando qualquer um dos separadores (/ ou _)
        numeros = re.split(r'[/_]', numeros_str)
        
        for numero in numeros:
            if numero:  # Garante que não adicionamos strings vazias
                os_encontradas.add(f"{numero}-{sufixo}")

    # --- ETAPA 2: Capturar todas as OS individuais no formato padrão ---
    # Este padrão encontrará tanto OS avulsas quanto as que já foram processadas na ETAPA 1.
    # O uso do 'set' cuida automaticamente da remoção de duplicatas.
    padrao_simples = re.compile(r'\b\d{2,4}-[A-Z0-9]{2,4}\b')
    os_encontradas.update(padrao_simples.findall(descricao))

    # É mais comum e prático retornar uma lista vazia do que None.
    return list(os_encontradas)