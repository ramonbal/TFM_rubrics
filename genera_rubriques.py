import os
import shutil
import re
import xlwings as xw
import time
import gc
import tempfile
from datetime import datetime

# operacions
# Si cal eliminar les rubriques existents . posar a True per a fer neteja, nomes borrara els fitxers de rubriques existents. despres posar a False
ELIMINAR_RUBRIQUES    = False  

CREAR_RUBRICA_ADVISOR = True  # Si cal crear la rubrica de l'advisor
CREAR_RUBRICA_COMITE  = True  # Si cal crear la rubrica del comitè
USE_LOCAL_TEMP        = True   # Treballar en directori temporal per evitar problemes OneDrive, deixar a True si es vol evitar problemes amb OneDrive

# Nom del fitxer de dades i plantilla
plantilla_comite  = "EvaluationGuidelinesCommitteEN.xlsx"
plantilla_advisor = "EvaluationGuidelinesAdvisorEN.xlsx"
fitxer_dades      = "committees.xlsx" #el fitxer te una fila de capçalera amb els noms de les columnes

 # Columnes on es troben les dades
col_estudiant = "A"  # Columna 1 (col_nom)
col_titol     = "B"  # Columna 2 (col_titol )
col_abstratc  = "C"  # Columna 3 (col_abstratc ) --- IGNORE ---
col_advisors  = "D"  # Columna 4 (col_advisors)  el fomrat es: "Advisor1 - Advisor2"
col_tribunal  = "E"  # Columna 5 (col_tribunal)  el format es: "President: nom_presi - secretari: nom_secre - vocal: nom_vocal"
col_advisor   = "F"  # Columna 6 (col_advisor1)  aqui es posa el nom de l'advisor principal despres de procesar
col_advisor2  = "G"  # Columna 7 (col_advisor2)  aqui es posa el nom de l'advisor secundari despres de procesar
col_president = "H"  # Columna 8 (col_president) aqui es posa el nom del president del comite despres de procesar
col_secretari = "I"  # Columna 9 (col_secretari) aqui es posa el nom del secretari del comite despres de procesar
col_vocal     = "J"  # Columna 10 (col_vocal)    aqui es posa el nom del vocal del comite despres de procesar
col_nota_adv  = "K"  # Columna 11 (col_nota_adv) aquesta columna es posa la nota de l'advisor, es crea un hipervincle a la rubrica de l'advisor
col_nota_com  = "L"  # Columna 12 (col_nota_com) aquesta columna es posa la nota del comite, es crea un hipervincle a la rubrica del comite

# Cel·les on es deixaran els valors del nom i el títol, president, secretari i vocal en la rubrica de comite
cella_estudiant_comite = "A4"
cella_titol_comite     = "A6"
cella_president_comite = "A78"  # Cella per al president, si cal
cella_secretari_comite = "C78"  # Cella per al secretari, si cal
cella_vocal_comite     = "G78"  # Cella per al vocal, si cal
cella_nota_adv_comite  = "H69"  # Cella per la nota del director
cella_nota_com_comite  = "H67"  # Cella per la nota del comitè - as requested by user

# Cel·les on es deixaran els valors del nom i el títol, advisors en la rubrica de advisor
cella_estudiant_advisor = "C3"
cella_titol_advisor     = "A5"
cella_advisor_advisor   = "C6"  # Cella per al advisor, si cal    
cella_signa_advisor     = "A26" # Cella per al nom advisor a la signatura
cella_nota_advisor      = "I22"  # Cella per la nota del director


# Canvia el directori de treball al directori del script
# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
onedrive_dir = script_dir  # Save original OneDrive directory



if ELIMINAR_RUBRIQUES:
    
    # Elimina tots els fitxers de rubrica advisor i comite de tots els subdirectoris
    for root, dirs, files in os.walk(script_dir):
        for file in files:
            # Only delete files in subdirectories, not in the main script directory
            if root != script_dir and file.endswith('.xlsx') and \
               (file.startswith(os.path.splitext(plantilla_advisor)[0]) or file.startswith(os.path.splitext(plantilla_comite)[0])):
                try:
                    os.remove(os.path.join(root, file))
                    print(f"🗑️ Eliminat: {file}")
                except Exception as e:
                    print(f"⚠️ Error eliminant {file}: {e}")
    exit()

if USE_LOCAL_TEMP:
    # Create a temporary working directory to avoid OneDrive issues
    temp_dir = tempfile.mkdtemp(prefix="rubrics_")
    print(f"🏠 Treballant en directori temporal: {temp_dir}")
    print(f"📁 Directori OneDrive original: {onedrive_dir}")
    
    # Copy template files to temp directory
    for template in [plantilla_advisor, plantilla_comite, fitxer_dades]:
        if os.path.exists(os.path.join(onedrive_dir, template)):
            shutil.copy2(os.path.join(onedrive_dir, template), temp_dir)
            print(f"📋 Copiat: {template}")
    
    os.chdir(temp_dir)
    script_dir = temp_dir
else:
    os.chdir(script_dir)  # Change the current working directory to the script's directory

print(f"📂 Directori de treball: {os.getcwd()}")

# Initialize xlwings app with better error handling
try:
    app = xw.App(visible=False)
    # Disable alerts to prevent Excel prompts
    app.display_alerts = False
    # Disable screen updating for better performance
    app.screen_updating = False
except Exception as e:
    print(f"Error initializing Excel application: {e}")
    raise

try:
    # Initialize variables to prevent NameError in exception handling
    wb_dades = None
    
    # Check if the data file exists
    if not os.path.exists(fitxer_dades):
        raise FileNotFoundError(f"El fitxer de dades '{fitxer_dades}' no existeix.")
    
    print("Verificant plantilles...")
    # Test if template files can be opened
    if CREAR_RUBRICA_ADVISOR:
        try:
            test_wb = app.books.open(os.path.abspath(plantilla_advisor))
            print(f"✓ Plantilla advisor accessible: {len(test_wb.sheets)} sheets")
            test_wb.close()
        except Exception as e:
            print(f"❌ Error amb plantilla advisor: {e}")
            raise
        
    if CREAR_RUBRICA_COMITE:
        try:
            test_wb = app.books.open(os.path.abspath(plantilla_comite))
            print(f"✓ Plantilla comitè accessible: {len(test_wb.sheets)} sheets")
            test_wb.close()
        except Exception as e:
            print(f"❌ Error amb plantilla comitè: {e}")
            raise
    
    # Open the data file
    wb_dades = app.books.open(os.path.abspath(fitxer_dades))
    ws_dades = wb_dades.sheets.active
    
    # Get max row for processing using xlwings
    max_row = ws_dades.range('A1').expand('down').last_cell.row
    
    print(f"Processant {max_row - 1} files de dades...")
    for fila in range(2, max_row + 1):  # Comença des de la fila 2
        # Defineix les cel·les de dades per a aquesta fila
        cella_estudiant_dades = f"{col_estudiant}{fila}"
        cella_titol_dades     = f"{col_titol}{fila}"
        cella_advisors_dades  = f"{col_advisors}{fila}"
        cella_advisor_dades   = f"{col_advisor}{fila}"
        cella_advisor2_dades  = f"{col_advisor2}{fila}"
        cella_tribunal_dades  = f"{col_tribunal}{fila}"
        cella_president_dades = f"{col_president}{fila}"
        cella_secretari_dades = f"{col_secretari}{fila}"
        cella_vocal_dades     = f"{col_vocal}{fila}"
        cella_nota_adv_dades  = f"{col_nota_adv}{fila}"
        cella_nota_com_dades  = f"{col_nota_com}{fila}"

        # Llegeix les dades 
        estudiant = str(ws_dades.range(cella_estudiant_dades).value).strip() if ws_dades.range(cella_estudiant_dades).value else ""
        titol     = str(ws_dades.range(cella_titol_dades).value).strip()     if ws_dades.range(cella_titol_dades).value else ""
        tribunal  = str(ws_dades.range(cella_tribunal_dades).value).strip()  if ws_dades.range(cella_tribunal_dades).value else ""
        advisors  = str(ws_dades.range(cella_advisors_dades).value).strip()  if ws_dades.range(cella_advisors_dades).value else ""

        # Salta files buides
        if not estudiant:
            continue
            
        subdirectori = estudiant.replace(" ", "_")

        # Crea el subdirectori si no existeix
        if not os.path.exists(subdirectori):
            os.makedirs(subdirectori)


        tribunal = [name.strip() for name in re.findall(r':\s*([^\(]+)\(', tribunal)]
        advisors = [name.strip() for name in re.findall(r'-\s*([^\(]+)\(', advisors)]
        
        president, secretari, vocal = tribunal
        advisor = advisors[0]
        advisor2 = advisors[1] if len(advisors) > 1 else ""
        advisors = ', '.join(advisors) if len(advisors) > 1 else advisors[0]

###################################################################################################
###########################  ADVISOR  #############################################################
        nom_fitxer_rubrica_advisor = f"{os.path.splitext(plantilla_advisor)[0]}_{subdirectori}.xlsx"
        nom_full_advisor = 'Rubrica advisor'
        if CREAR_RUBRICA_ADVISOR:
            # Nom del nou fitxer per al advisor
            ruta_sortida_advisor = os.path.join(subdirectori, nom_fitxer_rubrica_advisor)
            # Elimina fitxers existents si cal
            if ELIMINAR_RUBRIQUES and os.path.exists(ruta_sortida_advisor):
                try:
                    os.remove(ruta_sortida_advisor)
                except Exception as e:
                    print(f"⚠️ Error eliminant fitxer existent: {e}")
                    continue

            print(f"📝 Creant fitxer advisor: {ruta_sortida_advisor}")
            
            # Open a fresh copy of the advisor template for this student
            wb_advisor = app.books.open(os.path.abspath(plantilla_advisor))
            ws_advisor = wb_advisor.sheets.active
            nom_full_advisor = ws_advisor.name
            
            # Assigna els valors a les cel·les corresponents
            try:
                ws_advisor.range(cella_estudiant_advisor).value = estudiant
                ws_advisor.range(cella_titol_advisor).value     = titol
                ws_advisor.range(cella_advisor_advisor).value   = advisors
                ws_advisor.range(cella_signa_advisor).value     = advisor
            except Exception as e:
                print(f"Error assignant valors advisor: {e}")
                print(f"Cel·les: {cella_estudiant_advisor}, {cella_titol_advisor}, {cella_advisor_advisor}, {cella_signa_advisor}")
                wb_advisor.close()
                raise
            
            # Guarda el fitxer amb el nou nom
            try:
                wb_advisor.save(os.path.abspath(ruta_sortida_advisor))
            except Exception as e:
                print(f"Error guardant fitxer advisor: {e}")
                wb_advisor.close()
                raise
                
            # Close the workbook
            try:
                wb_advisor.close()
            except Exception as e:
                print(f"⚠️ Error tancant fitxer advisor: {e}")

            # crea l'hypervincle a la nota del director per a les dades
            hipervincle_nota_adv_dades  = f"'[{nom_fitxer_rubrica_advisor}]{nom_full_advisor}'!{cella_nota_advisor}"
        # crea l'hypervincle a la nota del director per al comite
        hipervincle_nota_adv_comite  = f"'[{nom_fitxer_rubrica_advisor}]{nom_full_advisor}'!{cella_nota_advisor}"


###################################################################################################
###########################  COMITE  #############################################################
        if CREAR_RUBRICA_COMITE:
            # Nom del nou fitxer per al comitè
            nom_fitxer_rubrica_comite = f"{os.path.splitext(plantilla_comite)[0]}_{subdirectori}.xlsx"
            ruta_sortida_comite = os.path.join(subdirectori, nom_fitxer_rubrica_comite)
            if ELIMINAR_RUBRIQUES and os.path.exists(ruta_sortida_comite):
                try:
                    os.remove(ruta_sortida_comite)
                except Exception as e:
                    print(f"⚠️ Error eliminant fitxer existent: {e}")
                    continue

            print(f"📝 Creant fitxer comitè: {ruta_sortida_comite}")
            
            # Open a fresh copy of the committee template for this student
            wb_comite = app.books.open(os.path.abspath(plantilla_comite))
            ws_comite = wb_comite.sheets.active
            nom_full_comite = ws_comite.name
                    
            # Assigna els valors a les cel·les corresponents
            try:
                ws_comite.range(cella_estudiant_comite).value = estudiant
                ws_comite.range(cella_titol_comite).value     = titol
                ws_comite.range(cella_president_comite).value = president
                ws_comite.range(cella_secretari_comite).value = secretari
                ws_comite.range(cella_vocal_comite).value     = vocal

                ws_comite.range(cella_nota_adv_comite).value = f"={hipervincle_nota_adv_comite}"
            except Exception as e:
                print(f"Error assignant valors comitè: {e}")
                print(f"Cel·les: {cella_estudiant_comite}, {cella_titol_comite}, {cella_president_comite}, {cella_secretari_comite}, {cella_vocal_comite}, {cella_nota_adv_comite}")
                wb_comite.close()
                raise

            # Guarda el fitxer amb el nou nom
            try:
                wb_comite.save(os.path.abspath(ruta_sortida_comite))
            except Exception as e:
                print(f"Error guardant fitxer comitè: {e}")
                wb_comite.close()
                raise
                
            # Close the workbook
            try:
                wb_comite.close()
            except Exception as e:
                print(f"⚠️ Error tancant fitxer comitè: {e}")

            # crea l'hypervincle a la nota del comite per a les dades
            hipervincle_nota_com_dades  = f"'[{nom_fitxer_rubrica_comite}]{nom_full_comite}'!{cella_nota_com_comite}"

###################################################################################################
###########################   DADES   #############################################################

        if CREAR_RUBRICA_ADVISOR:
            ws_dades.range(cella_advisor_dades).value   = advisor
            ws_dades.range(cella_advisor2_dades).value  = advisor2
            ws_dades.range(cella_nota_adv_dades).value = f"={hipervincle_nota_adv_dades}"

        if CREAR_RUBRICA_COMITE:
            ws_dades.range(cella_president_dades).value = president
            ws_dades.range(cella_secretari_dades).value = secretari
            ws_dades.range(cella_vocal_dades).value     = vocal
            ws_dades.range(cella_nota_com_dades).value = f"={hipervincle_nota_com_dades}"

    # Guarda els canvis al fitxer d'dades només si estem usant xlwings
    if CREAR_RUBRICA_COMITE or CREAR_RUBRICA_ADVISOR:
        try:
            wb_dades.save()
            print("✓ Hipervincles i dades guardats al fitxer de dades.")
        except Exception as e:
            print(f"⚠ Error guardant fitxer de dades: {e}")

    # Close data workbook
    try:
        if 'wb_dades' in locals() and wb_dades is not None:
            wb_dades.close()
            print("✓ Fitxer de dades tancat.")
    except Exception as e:
        print(f"⚠️ Error tancant fitxer de dades: {e}")

    # Copy results back to OneDrive if using temp directory
    if USE_LOCAL_TEMP and 'temp_dir' in locals():
        print(f"📤 Copiant resultats a OneDrive: {onedrive_dir}")
        
        try:
            copied_dirs = 0
            failed_dirs = []
            
            # Copy all generated directories and files back to OneDrive
            for item in os.listdir(temp_dir):
                src_path = os.path.join(temp_dir, item)
                dst_path = os.path.join(onedrive_dir, item)
                
                try:
                    if os.path.isdir(src_path):
                        # Create destination directory if it doesn't exist
                        os.makedirs(dst_path, exist_ok=True)
                        
                        # Copy all files from source to destination
                        for file_name in os.listdir(src_path):
                            src_file = os.path.join(src_path, file_name)
                            dst_file = os.path.join(dst_path, file_name)
                            if os.path.isfile(src_file):
                                shutil.copy2(src_file, dst_file)
                        print(f"📁 Copiat directori: {os.path.basename(dst_path)}")
                        copied_dirs += 1
                        
                    elif item == fitxer_dades:
                        # Copy updated data file
                        # Make a backup first
                        backup_path = os.path.join(onedrive_dir, f"{fitxer_dades}.backup")
                        if os.path.exists(dst_path):
                            shutil.copy2(dst_path, backup_path)
                        
                        shutil.copy2(src_path, dst_path)
                        print(f"📊 Copiat fitxer de dades actualitzat: {item}")
                        
                except Exception as e:
                    print(f"❌ Error copiant {item}: {e}")
                    failed_dirs.append(item)
            
            if failed_dirs:
                print(f"⚠️ No s'han pogut copiar {len(failed_dirs)} elements: {', '.join(failed_dirs)}")
                print(f"⚠️ Els fitxers es troben a: {temp_dir}")
                print("💡 Prova de copiar-los manualment o reinicia OneDrive.")
            else:
                print(f"✅ Tots els {copied_dirs} directoris copiats a OneDrive correctament!")
                
                # Clean up temp directory only if all copied successfully
                try:
                    shutil.rmtree(temp_dir)
                    print(f"🗑️ Directori temporal eliminat: {temp_dir}")
                except Exception as e:
                    print(f"⚠️ Error eliminant directori temporal: {e}")
                    print(f"📁 Directori temporal (buit): {temp_dir}")
                    print(f"💡 Per eliminar-lo manualment, executa:")
                    print(f'   Remove-Item -Path "{temp_dir}" -Recurse -Force')
                
        except Exception as e:
            print(f"❌ Error durant la còpia a OneDrive: {e}")
            print(f"⚠️ Els fitxers es troben a: {temp_dir}")
            print("💡 Pots copiar-los manualment des del directori temporal.")


except Exception as e:
    print(f"Error durant l'execució: {e}")
    import traceback
    traceback.print_exc()
    
    # If using temp directory, inform user where files are
    if USE_LOCAL_TEMP and 'temp_dir' in locals():
        print(f"⚠️ Els fitxers es troben a: {temp_dir}")
    
    # Try to close any open workbooks
    try:
        if 'wb_dades' in locals() and wb_dades is not None:
            wb_dades.close()
    except:
        pass
finally:
    # Close the Excel application
    try:
        if 'app' in locals() and app is not None:
            app.quit()
            time.sleep(0.2)  # Give time for Excel to close properly
            gc.collect()  # Force garbage collection
    except:
        pass

