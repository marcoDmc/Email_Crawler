import os
import shutil


def move_file_to_directory(directory_name, filepath):
    try:

        path = os.path.join(os.getcwd(), directory_name)

        if not os.path.exists(path):
            os.mkdir(path)
            print(f"Pasta '{directory_name}' criada com sucesso em {os.getcwd()}.")

        else:
            print(f"A pasta '{directory_name}' já existe em {os.getcwd()}.")

        shutil.move(filepath, path)

        print(f"Arquivo movido com sucesso para a pasta '{directory_name}'.")

    except Exception as e:
        print(f"Erro ao mover o arquivo para a pasta '{directory_name}': {e}")
