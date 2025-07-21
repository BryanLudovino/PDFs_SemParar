import PyInstaller.__main__
import os

if __name__ == '__main__':
    # Obtém o diretório onde o script está sendo executado
    workdir = os.path.abspath(os.path.dirname(__file__))
    
    # Define o caminho completo para a imagem
    image_path = os.path.join(workdir, 'abc.png')
    
    # Verifica se a imagem existe
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"A imagem 'abc.png' não foi encontrada no diretório '{workdir}'. Verifique se o arquivo está no lugar certo.")

    PyInstaller.__main__.run([
        '--name', 'Automação Sem Parar',
        'interface.py',
        '--onefile',
        '--windowed',
        f'--add-data={image_path}{os.pathsep}.'
    ]) 