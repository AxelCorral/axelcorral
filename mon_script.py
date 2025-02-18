import os
import pandas as pd
import sys
import subprocess
from tkinter import Tk, messagebox
from datetime import datetime


def merge_csv_files(input_folder, output_file):
    """
    Regroupe toutes les données des fichiers CSV dans un dossier donné en un seul fichier CSV,
    ajoute une colonne avec une date extraite du nom de fichier et une colonne concaténée avec Date et Reporting.

    :param input_folder: Chemin du dossier contenant les fichiers CSV.
    :param output_file: Chemin et nom du fichier CSV de sortie.
    """
    data_frames = []
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.csv'):
            file_path = os.path.join(input_folder, file_name)
            try:
                date_part = file_name.split('_')[-1].replace('.csv', '')
                df = pd.read_csv(file_path)
                df['Date'] = date_part
                if 'Reporting' in df.columns:
                    df['Date_Reporting'] = df['Date'] + '_' + df['Reporting'].astype(str)
                else:
                    raise ValueError("La colonne 'Reporting' est absente dans le fichier : " + file_name)
                data_frames.append(df)
                print(f"Fichier chargé : {file_name} avec la date {date_part}")
            except Exception as e:
                print(f"Erreur lors du chargement du fichier {file_name} : {e}")

    if data_frames:
        merged_df = pd.concat(data_frames, ignore_index=True)
        merged_df.to_csv(output_file, index=False)
        print(f"Données regroupées enregistrées dans : {output_file}")
    else:
        print("Aucun fichier CSV valide trouvé dans le dossier.")

def create_additional_tables(output_folder):
    """
    Crée trois tables supplémentaires : phase_table.csv, priority_table.csv et rating_table.csv
    dans le dossier de sortie spécifié.

    :param output_folder: Chemin du dossier de sortie.
    """
    phase_data = {
        "Phase": ["Preparation", "Design", "Realization", "Implementation"],
        "Order": [1, 2, 3, 4]
    }
    phase_table = pd.DataFrame(phase_data)
    phase_table.to_csv(os.path.join(output_folder, 'phase_table.csv'), index=False)

    priority_data = {
        "Priority": ["Low", "Medium", "High"],
        "Order": [1, 2, 3]
    }
    priority_table = pd.DataFrame(priority_data)
    priority_table.to_csv(os.path.join(output_folder, 'priority_table.csv'), index=False)

    rating_data = {
        "Rating": ["G", "R", "H", "Y", "P"]
    }
    rating_table = pd.DataFrame(rating_data)
    rating_table.to_csv(os.path.join(output_folder, 'rating_table.csv'), index=False)

    print("Les tables phase_table.csv, priority_table.csv et rating_table.csv ont été créées avec succès.")

def main():
    try:
        user_home = os.environ['USERPROFILE']
        selected_file = os.path.join(user_home, r"OneDrive - dsm-firmenich\Calendar\TIME TRACKING DMI.xlsx")
        output_folder = os.path.join(user_home, r"OneDrive - dsm-firmenich\PowerBI\Time_Tracking_Data\Data Files")
        os.makedirs(output_folder, exist_ok=True)

        file2 = os.path.join(output_folder, 'TIME_TRACKING_COMBINED.csv')
        file3 = os.path.join(output_folder, 'Book2.xlsx')

        # Appel à la fonction merge_csv_files
        merge_csv_files(output_folder, os.path.join(output_folder, 'project_table.csv'))

        sheet_2410 = pd.read_excel(selected_file, sheet_name='2410', header=1)
        sheet_2411 = pd.read_excel(selected_file, sheet_name='2411', header=1)
        sheet_2412 = pd.read_excel(selected_file, sheet_name='2412', header=1)
        sheet_2501 = pd.read_excel(selected_file, sheet_name='2501', header=1)
        sheet_2502 = pd.read_excel(selected_file, sheet_name='2502', header=1)
        sheet_data = pd.read_excel(selected_file, sheet_name='DATA', header=0)

        usecols_2410 = sheet_2410.columns[:40]
        usecols_2411 = sheet_2411.columns[:31]
        usecols_2412 = sheet_2412.columns[:32]
        usecols_2501 = sheet_2501.columns[:40]
        usecols_2502 = sheet_2502.columns[:40]


        sheet_2410_corrected = pd.read_excel(selected_file, sheet_name='2410', header=1, usecols=usecols_2410)
        sheet_2411_corrected = pd.read_excel(selected_file, sheet_name='2411', header=1, usecols=usecols_2411)
        sheet_2412_corrected = pd.read_excel(selected_file, sheet_name='2412', header=1, usecols=usecols_2412)
        sheet_2501_corrected = pd.read_excel(selected_file, sheet_name='2501', header=1, usecols=usecols_2501)
        sheet_2502_corrected = pd.read_excel(selected_file, sheet_name='2502', header=1, usecols=usecols_2502)

        with pd.ExcelWriter(file3) as writer:
            sheet_2410_corrected.to_excel(writer, sheet_name='2410', index=False, header=False)
            sheet_2411_corrected.to_excel(writer, sheet_name='2411', index=False, header=False)
            sheet_2412_corrected.to_excel(writer, sheet_name='2412', index=False, header=False)
            sheet_2501_corrected.to_excel(writer, sheet_name='2501', index=False, header=False)
            sheet_2502_corrected.to_excel(writer, sheet_name='2502', index=False, header=False)

            sheet_data.to_excel(writer, sheet_name='DATA', index=False)

        excel_data = pd.ExcelFile(selected_file)
        october_data = excel_data.parse('2410')
        november_data = excel_data.parse('2411')
        december_data = excel_data.parse('2412')
        january_data = excel_data.parse('2501')
        february_data = excel_data.parse('2502')


        def extract_month_data_corrected(month_data, month_name):
            month_data = month_data.loc[:38, :]
            days = month_data.iloc[1].values[1:]
            days = [int(day) if not pd.isna(day) else None for day in days]
            people_activities = month_data.iloc[3:]

            formatted_data = []
            for _, row in people_activities.iterrows():
                person = row[0]
                for j, activity in enumerate(row[1:]):
                    if pd.notna(activity) and days[j] is not None:
                        day = days[j]
                        date = f"{month_name}-{day:02d}"
                        formatted_data.append([person, day, activity, date])

            return pd.DataFrame(formatted_data, columns=["Personne", "Jour", "Activité", "Date"])

        october_transformed = extract_month_data_corrected(october_data, "2024-10")
        november_transformed = extract_month_data_corrected(november_data, "2024-11")
        december_transformed = extract_month_data_corrected(december_data, "2024-12")
        january_transformed = extract_month_data_corrected(january_data, "2025-01")
        february_transformed = extract_month_data_corrected(february_data, "2025-02")


        combined_data = pd.concat([october_transformed, november_transformed, december_transformed, january_transformed, february_transformed], ignore_index=True)
        combined_data.to_csv(file2, sep=';', index=False)

        df = pd.read_csv(file2, sep=';', encoding='utf-8')
        table_personnes = df[['Personne']].drop_duplicates().reset_index(drop=True)
        table_personnes['id_personne'] = range(1, len(table_personnes) + 1)
        table_personnes.to_csv(os.path.join(output_folder, 'table_personnes.csv'), index=False, sep=';', decimal=',')

        table_dates = df[['Date', 'Jour']].drop_duplicates().reset_index(drop=True)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        table_dates['Date'] = pd.to_datetime(table_dates['Date'], errors='coerce')
        table_dates['Jour_Num'] = table_dates['Date'].dt.day
        table_dates['Mois'] = table_dates['Date'].dt.month_name()
        table_dates['Année'] = table_dates['Date'].dt.year
        table_dates['id_date'] = range(1, len(table_dates) + 1)
        table_dates.to_csv(os.path.join(output_folder, 'table_dates.csv'), index=False, sep=';', decimal=',')

        table_activites = sheet_data[['DMI Activity', 'Sub Category', 'Category', 'Activity']].drop_duplicates().reset_index(drop=True)
        table_activites.rename(columns={
            'DMI Activity': 'Activité',
            'Category': 'Groupe_Activité',
            'Sub Category': 'Groupe_Activité_Détaillée',
            'Activity': 'Activity complete name'
        }, inplace=True)
        table_activites['id_activite'] = range(1, len(table_activites) + 1)
        table_activites.to_csv(os.path.join(output_folder, 'table_activites.csv'), index=False, sep=';', decimal=',')

        table_faits = df.merge(table_personnes, on='Personne', how='left') \
                        .merge(table_dates, on='Date', how='left') \
                        .merge(table_activites, left_on='Activité', right_on='Activité', how='left')
        table_faits['id_fait'] = range(1, len(table_faits) + 1)

        # Modifie Activité_Date pour inclure l'activité au lieu de START
        table_faits['Activité_Date'] = table_faits['Date'].dt.to_period('M').astype(str) + '_' + table_faits['Activité']

        table_faits = table_faits[['id_fait', 'id_personne', 'id_date', 'id_activite', 'Activité', 'Date', 'Activité_Date']]
        table_faits.to_csv(os.path.join(output_folder, 'table_faits.csv'), index=False, sep=';', decimal=',')

        # Création des nouvelles tables
        create_additional_tables(output_folder)

        root = Tk()
        root.withdraw()
        messagebox.showinfo("Succès", f"Tous les fichiers ont été créés avec succès et enregistrés dans :\n{output_folder}")

        power_bi_file = os.path.join(user_home, r"OneDrive - dsm-firmenich\PowerBI\Time_Tracking_Data\Activities Time Tracking.pbix")
        if messagebox.askyesno("Ouvrir Power BI", "Voulez-vous ouvrir le fichier Power BI ?"):
            subprocess.Popen([power_bi_file], shell=True)

    except Exception as e:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Erreur", f"Une erreur s'est produite :\n{str(e)}")

if __name__ == "__main__":
    main()

def merge_csv_files(input_folder, base_file, output_file):
    """
    Regroupe les données à partir d'un fichier CSV de base et de nouveaux fichiers dans un dossier donné en un seul fichier CSV,
    ajoute une colonne avec une date extraite du nom de fichier et une colonne concaténée avec Date et Reporting.
    
    :param input_folder: Chemin du dossier contenant les fichiers CSV.
    :param base_file: Chemin du fichier CSV de base.
    :param output_file: Chemin et nom du fichier CSV de sortie.
    """
    # Charger le fichier de base
    try:
        base_df = pd.read_csv(base_file)
        base_df = base_df.loc[:, ~base_df.columns.str.contains('^Unnamed')]

        base_df['Date'] = base_df.iloc[:, 10]  # Colonne K (index 10)
        
        # Convertir la date au format yyyy-MM si elle est sous forme yyyy-MM-dd
        base_df['Date'] = pd.to_datetime(base_df['Date'], errors='coerce').dt.strftime('%Y-%m')
        
        if 'Reporting' in base_df.columns:
            base_df['Date_Reporting'] = base_df['Date'].astype(str) + '_' + base_df['Reporting'].astype(str)
        else:
            raise ValueError("La colonne 'Reporting' est absente dans le fichier de base.")

        print(f"Fichier de base chargé : {base_file}")
    except Exception as e:
        print(f"Erreur lors du chargement du fichier de base : {e}")
        return

    # Liste pour stocker les DataFrames (incluant la base)
    data_frames = [base_df]

    # Parcourir tous les fichiers du dossier
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.csv') and file_name != os.path.basename(base_file):
            file_path = os.path.join(input_folder, file_name)
            try:
                # Extraire la date du nom de fichier et la convertir au format yyyy-MM
                date_part = file_name.split('_')[-1].replace('.csv', '')
                date_part = datetime.strptime(date_part, '%Y-%m-%d').strftime('%Y-%m')

                # Lire chaque fichier CSV
                df = pd.read_csv(file_path)
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

                # Ajouter la colonne de date
                df['Date'] = date_part
                
                # Créer une colonne concaténée avec Date et Reporting
                if 'Reporting' in df.columns:
                    df['Date_Reporting'] = df['Date'] + '_' + df['Reporting'].astype(str)
                else:
                    raise ValueError("La colonne 'Reporting' est absente dans le fichier : " + file_name)

                data_frames.append(df)
                print(f"Fichier chargé : {file_name} avec la date {date_part}")
            except Exception as e:
                print(f"Erreur lors du chargement du fichier {file_name} : {e}")

    if data_frames:
        # Combiner tous les DataFrames
        merged_df = pd.concat(data_frames, ignore_index=True)

        # Vider la première colonne
        if not merged_df.empty:
            merged_df.iloc[:, 0] = ''

        # Sauvegarder le DataFrame fusionné
        merged_df.to_csv(output_file, index=False)
        print(f"Données regroupées enregistrées dans : {output_file}")
    else:
        print("Aucun fichier CSV valide trouvé dans le dossier.")

# Exemple d'utilisation
input_folder = r"C:\Users\AXCO\OneDrive - dsm-firmenich\PowerBI\Time_Tracking_Data\Data Files\Project files"
base_file = r"C:\Users\AXCO\OneDrive - dsm-firmenich\PowerBI\Time_Tracking_Data\Data Files\Project files\project_table_2024_2025.csv"
output_file = r"C:\Users\AXCO\OneDrive - dsm-firmenich\PowerBI\Time_Tracking_Data\Data Files\project_table.csv"

merge_csv_files(input_folder, base_file, output_file)
