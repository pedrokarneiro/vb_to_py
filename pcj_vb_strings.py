#######################################################
# Funções de Manipulação de Strings no dialeto do VB6.
# Pedro Carneiro Jr.
# 2024-02-28
#######################################################

def Left(s, n):
    # Retorna os primeiros n caracteres da string s.
    return s[:n]

def Right(s, n):
    # Retorna os últimos n caracteres da string s.
    return s[-n:]

def Mid(s, start=1, length=None):
    # Retorna uma substring de comprimento 'length' começando da posição 'start'.
    # Gera um erro se 's' não for fornecido ou for uma string vazia.
    # Se 'start' não for fornecido, começa do início da string.
    # Se 'length' não for fornecido, retorna até o final da string.
    if s is None or not s:
        raise ValueError("O argumento 's' é obrigatório e não pode ser uma string vazia.")
    
    if length is None:
        return s[start - 1:]
    else:
        return s[start - 1:start - 1 + length]

def Len(s):
    # Retorna o comprimento da string s.
    # E você pode escolher a caixa da primeira letra da função. <:-D
    return len(s)

def InStr(s, substring, start=0):
    # Retorna a posição da primeira ocorrência de substring em string,
    # começando a busca a partir da posição 'start'.
    # Retorna 0 se não encontrar.
    return s.find(substring, start)

def Trim(s):
    # Remove espaços em branco do início e do final da string s.
    return s.strip()

def LTrim(s):
    # Remove os espaços em branco à esquerda de uma string.
    return s.lstrip()

def RTrim(s):
    # Remove os espaços em branco à direita de uma string.
    return s.rstrip()

def LCase(s):
    # Converte a string s para minúsculas.
    return s.lower()

def UCase(s):
    # Converte a string s para maiúsculas.
    return s.upper()

def Capitalize(s, sep=' '):
    # Verifica se o separador sep está presente na string.
    # Esse separador vai continuar na string após o processamento.
    if sep not in s:
        raise ValueError(f"O argumento 'sep' informado '{sep}' deve estar presente na string informada e não está.\n"
                         f"A string informada foi:\n"
                         f"\"{s}\"\n"
                         f"O separator 'sep' informado foi:\n"
                         f"\"{sep}\"")

    # Divide a string em palavras usando o separador
    words = s.split(sep)

    # Aplica a função capitalize a cada palavra
    capitalize_words = [word.capitalize() for word in words]

    # Junta as palavras novamente usando o separador e retorna o resultado
    return sep.join(capitalize_words)

def CamelCase(s, sep=' '):
    # Verifica se o separador sep está presente na string.
    # Esse separador vai ser retirado da string, ou seja, el vai ser concatenada sem ele.
    if sep not in s:
        raise ValueError(f"O argumento 'sep' informado '{sep}' deve estar presente na string informada e não está.\n"
                         f"A string informada foi:\n"
                         f"\"{s}\"\n"
                         f"O separator 'sep' informado foi:\n"
                         f"\"{sep}\"")

    # Divide a string em palavras usando o separador.
    words = s.split(sep)
    
    # Aplica a função capitalize a cada palavra menos na primeira.
    camel_case_words = [words[0].lower()] + [word.capitalize() for word in words[1:]]

    # Junta as palavras novamente sem o separador e retorna o resultado.
    return ''.join(camel_case_words)

def PascalCase(s, sep=' '):
    # Verifica se o separador sep está presente na string.
    # Esse separador vai ser retirado da string, ou seja, el vai ser concatenada sem ele.
    if sep not in s:
        raise ValueError(f"O argumento 'sep' informado '{sep}' deve estar presente na string informada e não está.\n"
                         f"A string informada foi:\n"
                         f"\"{s}\"\n"
                         f"O separator 'sep' informado foi:\n"
                         f"\"{sep}\"")

    # Divide a string em palavras usando o separador.
    words = s.split(sep)

    # Aplica a função capitalize a cada palavra.
    pascal_case_words = [word.capitalize() for word in words]

    # Junta as palavras novamente sem o separador e retorna o resultado.
    return ''.join(pascal_case_words)

def Replace(s, old, new, count=-1, start=1):
    # Substitui 'old' por 'new' em 's'.
    # Se 'count' for especificado, substitui 'count' ocorrências.
    # Se 'start' for especificado, começa a substituição a partir dessa posição.
    if not s:
        raise ValueError("O argumento 's' é obrigatório e não pode ser uma string vazia.")
    if not old:
        raise ValueError("O argumento 'old' é obrigatório e não pode ser uma string vazia.")
    if not new:
        raise ValueError("O argumento 'new' é obrigatório e não pode ser uma string vazia.")

    return s[:start - 1] + s[start - 1:].replace(old, new, count)

def InStrRev(s, substring, start=None):
    # Retorna a posição da última ocorrência de substring em string,
    # começando a busca a partir da posição 'start' (ou no final se não especificada).
    # Retorna 0 se não encontrar.

    # Verifica se os argumentos obrigatórios não são vazios
    if not s:
        raise ValueError("O argumento 's' é obrigatório e não pode ser uma string vazia.")
    if not substring:
        raise ValueError("O argumento 'substring' é obrigatório e não pode ser uma string vazia.")

    if start is None:
        start = len(s)
    index = s.rfind(substring, 0, start)
    return index + 1 if index != -1 else 0

def Space(n):
    # Retorna uma string com n espaços.
    return ' ' * n

def String(char, count):
    # Retorna uma string repetindo o caractere 'char' 'count' vezes.
    return char * count

def Split(s, delimitador=None, maxsplit=-1):
    # Divide a string em substrings usando um delimitador.
    return s.split(delimitador, maxsplit)

def StrComp(string1, string2, compare=0):
    """
        Compara duas strings e # Retorna um valor de comparação.
        0: comparação padrão,
        1: comparação sem diferenciar maiúsculas e minúsculas,
        2: comparação exata.
    """
    if compare == 0:
        return (string1 > string2) - (string1 < string2)
    elif compare == 1:
        return int(string1.lower() == string2.lower())
    elif compare == 2:
        return int(string1 == string2)
    else:
        raise ValueError("Invalid compare option. Use 0, 1, or 2.")

def StrReverse(s):
    # Inverte a ordem dos caracteres em uma string.
    return s[::-1]

def StrConv(s, conversion, localeID=0):
    # Converte uma string para maiúsculas (1) ou minúsculas (2).
    if conversion == 1:
        return s.upper()
    elif conversion == 2:
        return s.lower()
    else:
        raise ValueError("Opção de conversão inválida. Use 1 ou 2: maiúsculas (1) ou minúsculas (2)")

def Val(s):
    # Converte uma string para um número float.
    try:
        return float(s)
    except ValueError:
        return 0.0

def Chr(char_code):
    # Retorna um caractere pela sua posição Unicode.
    return chr(char_code)

def ChrW(char_code):
    # Retorna um caractere pela sua posição Unicode.
    return chr(char_code)

def Asc(char):
    # Retorna o valor Unicode de um caractere.
    return ord(char)

def AscW(char):
    # Retorna o valor Unicode de um caractere.
    return ord(char)

def PascalCase(s):
    words = s.split()
    pascal_case_words = [word.capitalize() for word in words[1:]]
    pascal_case_string = ' '.join([words[0]] + pascal_case_words)
    return pascal_case_string

if __name__ == '__main__':
    texto = "Exemplo de Texto"
    resultado = Left(texto, 4)
    print(resultado)  # Saída: "Exem"
