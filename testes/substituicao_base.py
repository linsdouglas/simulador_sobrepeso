import os
import shutil

def substituir_base_sap():
    print("[INÍCIO] Substituição da base SAP...")

    caminho_origem = os.path.join(os.environ["USERPROFILE"], "Downloads", "EXPORT.xlsx")
    pasta_destino = encontrar_pasta_onedrive_empresa()

    if not pasta_destino:
        raise FileNotFoundError("Pasta do OneDrive não localizada.")
    
    caminho_destino = os.path.join(pasta_destino, "base_sap.xlsx")

    if not os.path.exists(caminho_origem):
        print("[ERRO] Arquivo 'EXPORT.xlsx' não encontrado na pasta Downloads.")
        return

    if os.path.exists(caminho_destino):
        print("Apagando base antiga do OneDrive...")
        os.remove(caminho_destino)

    print("Copiando 'EXPORT.xlsx' para o OneDrive como 'base_sap.xlsx'...")
    shutil.copyfile(caminho_origem, caminho_destino)

    print("[SUCESSO] Substituição concluída. Novo arquivo salvo em:")
    print(caminho_destino)

def encontrar_pasta_onedrive_empresa():
    print("[INFO] Procurando pasta do OneDrive sincronizada com SharePoint...")
    user_dir = os.environ["USERPROFILE"]
    possiveis = os.listdir(user_dir)
    for nome in possiveis:
        if "DIAS BRANCO" in nome.upper():
            caminho_completo = os.path.join(user_dir, nome)
            if os.path.isdir(caminho_completo) and "Gestão de Estoque - Documentos" in os.listdir(caminho_completo):
                print(f"[OK] Pasta encontrada: {caminho_completo}")
                return os.path.join(caminho_completo, "Gestão de Estoque - Documentos")
    print("[ERRO] Pasta não encontrada.")
    return None

if __name__ == "__main__":
    substituir_base_sap()
