import streamlit as st
import pandas as pd
import numpy as np
from urllib.parse import urlparse
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import zipfile
from datetime import datetime

def main():
    st.set_page_config(
        page_title="Audit Sémantique - Gap Content SEO",
        page_icon="🔍",
        layout="wide"
    )
    
    st.title("Outil d'Audit Sémantique - Gap Content SEO")
    st.markdown("---")
    
    # Sidebar pour la configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Section 1: Type de source de données
        st.subheader("1. Source des données")
        data_source = st.selectbox(
            "Type de fichier d'export",
            ["Semrush", "Ahrefs", "Custom"]
        )
        
            if data_source == "Custom":
                st.subheader("2. Mapping des colonnes")
                col_keyword = st.text_input("Nom colonne Mot-clé", "Keyword")
                col_domain = st.text_input("Nom colonne Domaine", "Domain")
                col_position = st.text_input("Nom colonne Position", "Position")
                col_volume = st.text_input("Nom colonne Volume de recherche", "Search Volume")
                col_difficulty = st.text_input("Nom colonne Difficulté", "Keyword Difficulty")
                col_intent = st.text_input("Nom colonne Intention", "Keyword Intents")
                col_url = st.text_input("Nom colonne URL", "URL")
            else:
                # Configuration prédéfinie pour Semrush/Ahrefs
                if data_source == "Semrush":
                    col_mapping = {
                        'keyword': 'Keyword',
                        'domain': 'URL',  # On extraira le domaine de l'URL
                        'position': 'Position',
                        'volume': 'Search Volume',
                        'difficulty': 'Keyword Difficulty',
                        'intent': 'Keyword Intents',
                        'url': 'URL'
                    }
                else:  # Ahrefs
                    col_mapping = {
                        'keyword': 'Keyword',
                        'domain': 'Domain',
                        'position': 'Position',
                        'volume': 'Volume',
                        'difficulty': 'KD',
                        'intent': 'Intent',
                        'url': 'URL'
                    }
            
            # Section 2/3: Critères de filtrage
            st.subheader("2. Critères Gap Content")
            min_competitors = st.selectbox(
                "Nombre minimum de concurrents positionnés",
                [1, 2, 3]
            )
            
            max_position = st.selectbox(
                "Position maximum des concurrents",
                [10, 20, 50]
            )
            
            # Filtres supplémentaires
            st.subheader("3. Filtres supplémentaires")
            min_volume = st.number_input("Volume de recherche minimum", min_value=0, value=0)
            max_difficulty = st.number_input("Difficulté maximum", min_value=0, max_value=100, value=100)

    # Zone principale
    st.header("📁 Import des fichiers")
    
    # Upload des fichiers
    uploaded_files = st.file_uploader(
        "Téléchargez vos fichiers CSV/Excel (le premier sera considéré comme votre domaine principal)",
        accept_multiple_files=True,
        type=['csv', 'xlsx', 'xls']
    )
    
    # Identification du domaine principal
    if uploaded_files:
        st.subheader("Identification du domaine principal")
        
        # Option 1: Premier fichier par défaut
        main_domain_option = st.radio(
            "Comment identifier votre domaine principal ?",
            ["Premier fichier téléchargé", "Sélection manuelle"]
        )
        
        if main_domain_option == "Sélection manuelle":
            main_domain = st.text_input("Saisissez votre domaine principal (ex: www.monsite.com)")
        else:
            main_domain = None
    
    # Bouton d'analyse
    if uploaded_files and st.button("Lancer l'analyse", type="primary"):
        with st.spinner("Analyse en cours..."):
            try:
                # Traitement des fichiers
                all_data = process_files(uploaded_files, data_source, locals())
                
                if all_data is not None and not all_data.empty:
                    # Identification du domaine principal
                    if main_domain_option == "Premier fichier téléchargé":
                        first_file_data = load_file(uploaded_files[0], data_source, locals())
                        if not first_file_data.empty:
                            main_domain = extract_domain_from_data(first_file_data)
                    
                    # Analyse du gap content
                    gap_analysis = perform_gap_analysis(
                        all_data, 
                        main_domain, 
                        min_competitors, 
                        max_position,
                        min_volume,
                        max_difficulty
                    )
                    
                    if gap_analysis['gap_content'].empty:
                        st.warning("Aucune opportunité de gap content trouvée avec ces critères.")
                    else:
                        # Génération du fichier Excel
                        excel_file = generate_excel_report(gap_analysis, main_domain)
                        
                        # Affichage d'un résumé simple
                        st.success(f"Analyse terminée ! {len(gap_analysis['gap_content'])} opportunités trouvées.")
                        
                        if not gap_analysis['gap_content'].empty:
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Total opportunités", len(gap_analysis['gap_content']))
                            with col2:
                                avg_volume = gap_analysis['gap_content']['volume'].mean()
                                st.metric("Volume moyen", f"{avg_volume:,.0f}")
                            with col3:
                                total_volume = gap_analysis['gap_content']['volume'].sum()
                                st.metric("Volume total", f"{total_volume:,.0f}")
                        
                        # Bouton de téléchargement
                        st.download_button(
                            label="Télécharger le rapport Excel",
                            data=excel_file,
                            file_name=f"gap_content_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("Impossible de traiter les fichiers téléchargés.")
                    
            except Exception as e:
                st.error(f"Erreur lors de l'analyse : {str(e)}")


def extract_domain(url):
    """Extrait le domaine d'une URL"""
    try:
        if pd.isna(url) or url == '':
            return ''
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        parsed = urlparse(url)
        return parsed.netloc.lower()
    except:
        return ''


def load_file(file, data_source, config):
    """Charge un fichier CSV ou Excel et normalise les colonnes"""
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        
        # Normalisation des colonnes selon la source
        if data_source == "Semrush":
            if 'Keyword' in df.columns:
                df['domain'] = df['URL'].apply(extract_domain)
                df = df.rename(columns={
                    'Keyword': 'keyword',
                    'Position': 'position',
                    'Search Volume': 'volume',
                    'Keyword Difficulty': 'difficulty',
                    'Keyword Intents': 'intent',
                    'URL': 'url'
                })
        elif data_source == "Ahrefs":
            # Configuration à adapter selon le format Ahrefs réel
            df['domain'] = df.get('URL', df.get('Domain', '')).apply(extract_domain)
            # À compléter selon les colonnes Ahrefs
        else:  # Custom
            df['domain'] = df[config.get('col_domain', 'URL')].apply(extract_domain)
            df = df.rename(columns={
                config.get('col_keyword', 'Keyword'): 'keyword',
                config.get('col_position', 'Position'): 'position',
                config.get('col_volume', 'Search Volume'): 'volume',
                config.get('col_difficulty', 'Keyword Difficulty'): 'difficulty',
                config.get('col_intent', 'Keyword Intents'): 'intent',
                config.get('col_url', 'URL'): 'url'
            })
        
        # Nettoyage des données
        df = df.dropna(subset=['keyword', 'domain'])
        df['position'] = pd.to_numeric(df['position'], errors='coerce')
        df['volume'] = pd.to_numeric(df['volume'], errors='coerce')
        df['difficulty'] = pd.to_numeric(df['difficulty'], errors='coerce')
        
        # Ajout du nom du fichier
        df['source_file'] = file.name
        
        return df
        
    except Exception as e:
        st.error(f"Erreur lors du chargement de {file.name}: {str(e)}")
        return pd.DataFrame()


def process_files(files, data_source, config):
    """Traite tous les fichiers uploadés"""
    all_dataframes = []
    
    for file in files:
        df = load_file(file, data_source, config)
        if not df.empty:
            all_dataframes.append(df)
    
    if all_dataframes:
        return pd.concat(all_dataframes, ignore_index=True)
    else:
        return pd.DataFrame()


def extract_domain_from_data(df):
    """Extrait le domaine principal du premier fichier"""
    if 'domain' in df.columns and not df['domain'].empty:
        # Prendre le domaine le plus fréquent
        return df['domain'].value_counts().index[0]
    return None


def perform_gap_analysis(data, main_domain, min_competitors, max_position, min_volume, max_difficulty):
    """Effectue l'analyse du gap content"""
    
    # Filtrage des positions valides
    data = data[(data['position'] <= max_position) & (data['position'] > 0)]
    
    # Filtrage par volume et difficulté
    if min_volume > 0:
        data = data[data['volume'] >= min_volume]
    if max_difficulty < 100:
        data = data[data['difficulty'] <= max_difficulty]
    
    # Groupement par mot-clé
    keyword_analysis = []
    
    for keyword, group in data.groupby('keyword'):
        # Informations du mot-clé
        volume = group['volume'].iloc[0] if not group['volume'].isna().all() else 0
        difficulty = group['difficulty'].iloc[0] if not group['difficulty'].isna().all() else 0
        intent = group['intent'].iloc[0] if 'intent' in group.columns else ''
        
        # Analyse des domaines positionnés
        positioned_domains = group[group['position'] <= max_position]
        unique_domains = positioned_domains['domain'].unique()
        
        # Vérifier si le domaine principal est présent
        main_domain_present = main_domain in unique_domains if main_domain else False
        competitor_count = len(unique_domains) - (1 if main_domain_present else 0)
        
        # Critère de gap content : assez de concurrents positionnés MAIS domaine principal absent
        if competitor_count >= min_competitors and not main_domain_present:
            
            keyword_data = {
                'keyword': keyword,
                'volume': volume,
                'difficulty': difficulty,
                'intent': intent,
                'competitor_count': competitor_count,
                'total_domains': len(unique_domains)
            }
            
            # Ajouter les positions et URLs de chaque domaine
            for domain in unique_domains:
                domain_data = positioned_domains[positioned_domains['domain'] == domain]
                best_position = domain_data['position'].min()
                best_url = domain_data[domain_data['position'] == best_position]['url'].iloc[0]
                
                keyword_data[f'{domain}_position'] = best_position
                keyword_data[f'{domain}_url'] = best_url
            
            keyword_analysis.append(keyword_data)
    
    gap_content_df = pd.DataFrame(keyword_analysis)
    
    # Création du rapport par domaine
    domain_reports = {}
    for domain in data['domain'].unique():
        if domain and domain != main_domain:
            domain_data = data[data['domain'] == domain].copy()
            domain_data = domain_data.sort_values(['position', 'volume'], ascending=[True, False])
            domain_reports[domain] = domain_data[['keyword', 'volume', 'position', 'url']].reset_index(drop=True)
    
    return {
        'gap_content': gap_content_df,
        'domain_reports': domain_reports,
        'main_domain': main_domain,
        'all_data': data
    }


def display_results(analysis):
    """Affiche les résultats de l'analyse"""
    pass  # Fonction supprimée - seuls les résultats Excel sont nécessaires


def generate_excel_report(analysis, main_domain):
    """Génère le rapport Excel avec mise en forme"""
    
    output = io.BytesIO()
    workbook = Workbook()
    
    # Suppression de la feuille par défaut
    workbook.remove(workbook.active)
    
    # Styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center")
    
    # 1. Onglet Gap Content Analysis
    if not analysis['gap_content'].empty:
        ws_gap = workbook.create_sheet("Gap Content Analysis")
        gap_df = analysis['gap_content'].copy()
        
        # Réorganiser les colonnes
        base_cols = ['keyword', 'volume', 'difficulty', 'intent', 'competitor_count']
        domain_cols = [col for col in gap_df.columns if '_position' in col or '_url' in col]
        all_cols = base_cols + sorted(domain_cols)
        
        gap_df_display = gap_df[all_cols].copy()
        
        # Renommer les colonnes pour l'affichage
        column_mapping = {
            'keyword': 'Mot-clé',
            'volume': 'Volume de recherche',
            'difficulty': 'Difficulté concurrentielle',
            'intent': 'Intention de recherche',
            'competitor_count': 'Nombre de sites présents'
        }
        
        # Ajouter les noms de domaines dans le mapping
        for col in domain_cols:
            if '_position' in col:
                domain = col.replace('_position', '')
                column_mapping[col] = f'{domain} (Position)'
            elif '_url' in col:
                domain = col.replace('_url', '')
                column_mapping[col] = f'{domain} (URL)'
        
        gap_df_display = gap_df_display.rename(columns=column_mapping)
        
        # Écriture des données
        for r in dataframe_to_rows(gap_df_display, index=False, header=True):
            ws_gap.append(r)
        
        # Mise en forme des en-têtes
        for cell in ws_gap[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
        
        # Ajustement des largeurs de colonnes
        for column in ws_gap.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws_gap.column_dimensions[column_letter].width = adjusted_width
    
    # 2. Onglets pour chaque domaine concurrent
    for domain, domain_data in analysis['domain_reports'].items():
        if not domain_data.empty:
            # Créer un nom d'onglet valide (max 31 caractères)
            sheet_name = domain.replace('www.', '').replace('.com', '').replace('.fr', '')[:31]
            ws_domain = workbook.create_sheet(sheet_name)
            
            # Ligne de titre avec le nom du domaine
            ws_domain.append([domain])
            ws_domain['A1'].font = Font(size=18, bold=True)
            ws_domain.merge_cells('A1:D1')
            
            # Ligne vide
            ws_domain.append([])
            
            # En-têtes
            headers = ['Mot-clé', 'Volume de recherche', 'Position', 'URL']
            ws_domain.append(headers)
            
            # Mise en forme des en-têtes
            for cell in ws_domain[3]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_alignment
            
            # Données triées par position puis volume
            sorted_data = domain_data.sort_values(['position', 'volume'], ascending=[True, False])
            
            # Écriture des données
            for _, row in sorted_data.iterrows():
                ws_domain.append([row['keyword'], row['volume'], row['position'], row['url']])
            
            # Mise en forme conditionnelle pour les volumes
            if len(sorted_data) > 0:
                min_volume = sorted_data['volume'].min()
                max_volume = sorted_data['volume'].max()
                
                # Coloration conditionnelle des volumes (colonne B)
                for row_num in range(4, len(sorted_data) + 4):
                    cell = ws_domain[f'B{row_num}']
                    volume = cell.value
                    if volume and max_volume > min_volume:
                        # Gradient du blanc au vert foncé
                        intensity = (volume - min_volume) / (max_volume - min_volume)
                        green_value = int(255 - (intensity * 100))  # De 255 (blanc) à 155 (vert clair)
                        color = f"FF{green_value:02X}FF{green_value:02X}"
                        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            
            # Ajustement des largeurs
            for column in ws_domain.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws_domain.column_dimensions[column_letter].width = adjusted_width
    
    # Sauvegarde
    workbook.save(output)
    output.seek(0)
    
    return output.getvalue()


if __name__ == "__main__":
    main()
