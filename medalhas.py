import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class OlympicMedalsViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Visualizador de Medalhas Olímpicas")
        self.root.geometry("950x650")
        
        self.tabela_visivel = False

        self.df = self.carregar_dados()
        
        # Lista dos anos e países disponíveis
        self.anos_disponiveis = sorted(self.df['year'].unique()) if not self.df.empty else []
        self.paises_disponiveis = sorted(self.df['country'].unique()) if not self.df.empty else []
        
        # Cria a interface gráfica
        self.criar_interface()
    
    # Carrega os dados do arquivo Excel
    def carregar_dados(self):
        try:
            return pd.read_excel("world_olympedia_olympics_game_medal_tally.xlsx")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o Excel: {e}")
            return pd.DataFrame()
    
    def criar_interface(self):

        # Consulta 1: Top 10 países com mais medalhas em um ano
        self.consulta1_frame = ttk.LabelFrame(self.root, text="1. Top 10 países com mais medalhas em determinado ano")
        self.consulta1_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.filtro_ano_frame = ttk.Frame(self.consulta1_frame)
        self.filtro_ano_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.filtro_ano_frame, text="Selecione o ano:").pack(side="left")
        self.combo_ano = ttk.Combobox(self.filtro_ano_frame, values=self.anos_disponiveis, state="readonly", width=10)
        self.combo_ano.pack(side="left", padx=5)
        if self.anos_disponiveis:
            self.combo_ano.set(self.anos_disponiveis[0])
        ttk.Button(self.filtro_ano_frame, text="Buscar", command=self.filtrar_por_ano).pack(side="left", padx=5)
        self.label_status = ttk.Label(self.filtro_ano_frame, text="Selecione o ano e clique em 'Buscar'")
        self.label_status.pack(side="left", padx=20)
        
        # Consulta 2: Total de medalhas de um país
        self.consulta2_frame = ttk.LabelFrame(self.root, text="2. Total de medalhas de um país específico")
        self.consulta2_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.filtro_pais_frame = ttk.Frame(self.consulta2_frame)
        self.filtro_pais_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.filtro_pais_frame, text="Selecione o país:").pack(side="left")
        self.combo_pais = ttk.Combobox(self.filtro_pais_frame, values=self.paises_disponiveis, state="readonly", width=30)
        self.combo_pais.pack(side="left", padx=5)
        if self.paises_disponiveis:
            self.combo_pais.set(self.paises_disponiveis[0])
        ttk.Button(self.filtro_pais_frame, text="Buscar", command=self.filtrar_por_pais).pack(side="left", padx=5)
        self.label_status_pais = ttk.Label(self.filtro_pais_frame, text="Selecione o país e clique em 'Buscar'")
        self.label_status_pais.pack(side="left", padx=20)

        # Consulta 3: Top 10 países com mais ouros
        self.consulta3_frame = ttk.LabelFrame(self.root, text="3. Top 10 países com mais medalhas de ouro na história")
        self.consulta3_frame.pack(fill=tk.X, padx=10, pady=5)

        self.filtro_ouro_frame = ttk.Frame(self.consulta3_frame)
        self.filtro_ouro_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(self.filtro_ouro_frame, text="Buscar", command=self.buscar_pais_mais_ouro).pack(side="left", padx=5)
        self.label_status_ouro = ttk.Label(self.filtro_ouro_frame, text="Clique em 'Buscar' para ver os países com mais medalhas de ouro")
        self.label_status_ouro.pack(side="left", padx=20)

        # Consulta 4: Países sem medalha de ouro
        self.consulta4_frame = ttk.LabelFrame(self.root, text="4. Países que nunca ganharam medalhas de ouro")
        self.consulta4_frame.pack(fill=tk.X, padx=10, pady=5)

        self.filtro_sem_ouro_frame = ttk.Frame(self.consulta4_frame)
        self.filtro_sem_ouro_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(self.filtro_sem_ouro_frame, text="Buscar", command=self.buscar_sem_ouro).pack(side="left", padx=5)
        self.label_status_sem_ouro = ttk.Label(self.filtro_sem_ouro_frame, text="Clique em 'Buscar' para ver os países que nunca ganharam medalhas de ouro")
        self.label_status_sem_ouro.pack(side="left", padx=20)

        # Consulta 5: Edição mais competitiva
        self.consulta5_frame = ttk.LabelFrame(self.root, text="5. Edição mais competitiva (países com mais medalhas)")
        self.consulta5_frame.pack(fill=tk.X, padx=10, pady=5)

        self.filtro_competitiva_frame = ttk.Frame(self.consulta5_frame)
        self.filtro_competitiva_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(self.filtro_competitiva_frame, text="Buscar", command=self.buscar_edicao_mais_competitiva).pack(side="left", padx=5)
        self.label_status_competitiva = ttk.Label(self.filtro_competitiva_frame, text="Clique em 'Buscar' para ver a edição com mais países medalhistas")
        self.label_status_competitiva.pack(side="left", padx=20)

        # Consulta 6: Países com medalhas em todas as edições
        self.consulta6_frame = ttk.LabelFrame(self.root, text="6. Países que ganharam medalhas em todas as edições")
        self.consulta6_frame.pack(fill=tk.X, padx=10, pady=5)

        self.filtro_todas_edicoes_frame = ttk.Frame(self.consulta6_frame)
        self.filtro_todas_edicoes_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(self.filtro_todas_edicoes_frame, text="Buscar", command=self.buscar_paises_todas_edicoes).pack(side="left", padx=5)
        self.label_status_todas_edicoes = ttk.Label(self.filtro_todas_edicoes_frame, text="Clique em 'Buscar' para ver os países que ganharam medalhas em todas as edições")
        self.label_status_todas_edicoes.pack(side="left", padx=20)

        # Botão para exibir/ocultar tabela completa
        self.botao_mostrar = ttk.Button(self.root, text="Mostrar Tabela Completa", command=self.alternar_tabela)
        self.botao_mostrar.pack(pady=10)

        # Frame da tabela de resultados
        self.frame_tabela = ttk.Frame(self.root)
        self.criar_tabela()
    
    # Criação da Treeview para exibir os dados
    def criar_tabela(self):
        self.colunas = list(self.df.columns) if not self.df.empty else []
        vsb = ttk.Scrollbar(self.frame_tabela, orient="vertical")
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(self.frame_tabela, orient="horizontal")
        hsb.pack(side="bottom", fill="x")
        self.tree = ttk.Treeview(self.frame_tabela, columns=self.colunas, show='headings', yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        for col in self.colunas:
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=120, anchor='center')
    
    # Atualiza os dados na Treeview
    def atualizar_treeview(self, df_novo):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if list(df_novo.columns) != self.colunas:
            self.tree["columns"] = list(df_novo.columns)
            for col in df_novo.columns:
                self.tree.heading(col, text=col.capitalize())
                self.tree.column(col, width=120, anchor='center')
            self.colunas = list(df_novo.columns)
        for _, row in df_novo.iterrows():
            self.tree.insert("", tk.END, values=list(row))
    
    # Consulta 1: Filtra os dados por ano
    def filtrar_por_ano(self):
        if not self.combo_ano.get():
            messagebox.showwarning("Atenção", "Selecione um ano.")
            return
        ano = int(self.combo_ano.get())
        df_filtrado = self.df[self.df['year'] == ano]
        if df_filtrado.empty:
            self.atualizar_treeview(pd.DataFrame())
            self.label_status.config(text="Nenhum dado encontrado para o ano informado.")
            return
        resumo = df_filtrado.groupby(['country', 'country_noc']).agg({
            'gold': 'sum',
            'silver': 'sum',
            'bronze': 'sum',
            'total': 'sum'
        }).reset_index()
        top10 = resumo.sort_values(by='total', ascending=False).head(10)
        if not self.tabela_visivel:
            self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            self.botao_mostrar.config(text="Ocultar Tabela")
            self.tabela_visivel = True
        self.atualizar_treeview(top10)
        self.label_status.config(text=f"Top 10 países em {ano}")
    
    # Consulta 2: Filtra os dados por país
    def filtrar_por_pais(self):
        if not self.combo_pais.get():
            messagebox.showwarning("Atenção", "Selecione um país.")
            return
        pais = self.combo_pais.get()
        df_filtrado = self.df[self.df['country'] == pais]
        if df_filtrado.empty:
            self.atualizar_treeview(pd.DataFrame())
            self.label_status_pais.config(text="Nenhum dado encontrado para o país informado.")
            return
        resumo = df_filtrado.groupby(['year', 'country', 'country_noc']).agg({
            'gold': 'sum',
            'silver': 'sum',
            'bronze': 'sum',
            'total': 'sum'
        }).reset_index().sort_values(by='year', ascending=False)
        if not self.tabela_visivel:
            self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            self.botao_mostrar.config(text="Ocultar Tabela")
            self.tabela_visivel = True
        self.atualizar_treeview(resumo)
        self.label_status_pais.config(text=f"Medalhas de {pais} em todas as olimpíadas")
    
    # Consulta 3: Top países com mais ouros
    def buscar_pais_mais_ouro(self):
        if self.df.empty:
            self.label_status_ouro.config(text="Arquivo não carregado ou vazio.")
            return
        try:
            resumo = self.df.groupby(['country', 'country_noc']).agg({
                'gold': 'sum',
                'silver': 'sum',
                'bronze': 'sum',
                'total': 'sum'
            }).reset_index()
            top_ouro = resumo.sort_values(by='gold', ascending=False).head(10)
            if not self.tabela_visivel:
                self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                self.botao_mostrar.config(text="Ocultar Tabela")
                self.tabela_visivel = True
            self.atualizar_treeview(top_ouro)
            self.label_status_ouro.config(text="Top 10 países com mais medalhas de ouro exibido abaixo")
        except Exception as e:
            self.label_status_ouro.config(text=f"Erro: {e}")

    # Consulta 4: Países que nunca ganharam ouro
    def buscar_sem_ouro(self):
        if self.df.empty:
            self.label_status_sem_ouro.config(text="Arquivo não carregado ou vazio.")
            return
        try:
            resumo = self.df.groupby(['country', 'country_noc']).agg({
                'gold': 'sum',
                'silver': 'sum',
                'bronze': 'sum',
                'total': 'sum'
            }).reset_index()
            sem_ouro = resumo[resumo['gold'] == 0].sort_values(by='total', ascending=False)
            if sem_ouro.empty:
                self.label_status_sem_ouro.config(text="Todos os países já ganharam ouro.")
                return
            if not self.tabela_visivel:
                self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                self.botao_mostrar.config(text="Ocultar Tabela")
                self.tabela_visivel = True
            self.atualizar_treeview(sem_ouro)
            self.label_status_sem_ouro.config(text=f"{len(sem_ouro)} países que nunca ganharam medalha de ouro")
        except Exception as e:
            self.label_status_sem_ouro.config(text=f"Erro: {e}")

    # Consulta 5: Edição mais competitiva
    def buscar_edicao_mais_competitiva(self):
        if self.df.empty:
            self.label_status_competitiva.config(text="Arquivo não carregado ou vazio.")
            return
        try:
            df_filtrado = self.df[self.df['total'] > 0]
            contagem = df_filtrado.groupby('year')['country'].nunique().reset_index(name='paises_com_medalha')
            ano_max = contagem.sort_values(by='paises_com_medalha', ascending=False).iloc[0]['year']
            paises_ano = df_filtrado[df_filtrado['year'] == ano_max][['country', 'country_noc', 'gold', 'silver', 'bronze', 'total']]
            resumo = paises_ano.groupby(['country', 'country_noc']).agg({
                'gold': 'sum',
                'silver': 'sum',
                'bronze': 'sum',
                'total': 'sum'
            }).reset_index().sort_values(by='total', ascending=False)
            if not self.tabela_visivel:
                self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                self.botao_mostrar.config(text="Ocultar Tabela")
                self.tabela_visivel = True
            self.atualizar_treeview(resumo)
            self.label_status_competitiva.config(text=f"Os {len(resumo)} países que ganharam medalhas em {int(ano_max)} (edição mais competitiva)")
        except Exception as e:
            self.label_status_competitiva.config(text=f"Erro: {e}")

    # Consulta 6: Países com medalhas em todas as edições
    def buscar_paises_todas_edicoes(self):
        if self.df.empty:
            self.label_status_todas_edicoes.config(text="Arquivo não carregado ou vazio.")
            return
        try:
            anos = set(self.df['year'].unique())
            df_filtrado = self.df[self.df['total'] > 0]
            pais_edicoes = df_filtrado.groupby('country')['year'].apply(set).reset_index()
            paises_todas = pais_edicoes[pais_edicoes['year'].apply(lambda x: anos.issubset(x))]
            paises_todas = paises_todas[['country']].sort_values(by='country')
            if paises_todas.empty:
                self.label_status_todas_edicoes.config(text="Nenhum país ganhou medalhas em todas as edições.")
                return
            if not self.tabela_visivel:
                self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                self.botao_mostrar.config(text="Ocultar Tabela")
                self.tabela_visivel = True
            self.atualizar_treeview(paises_todas)
            self.label_status_todas_edicoes.config(text=f"Apenas {len(paises_todas)} países ganharam medalhas em todas as edições")
        except Exception as e:
            self.label_status_todas_edicoes.config(text=f"Erro: {e}")

    # Alterna a visibilidade da tabela completa
    def alternar_tabela(self):
        if self.tabela_visivel:
            self.frame_tabela.pack_forget()
            self.botao_mostrar.config(text="Mostrar Tabela Completa")
            self.tabela_visivel = False
        else:
            self.frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            self.atualizar_treeview(self.df)
            self.label_status.config(text=f"Exibindo todas as {len(self.df)} linhas")
            self.botao_mostrar.config(text="Ocultar Tabela")
            self.tabela_visivel = True

# Execução da aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = OlympicMedalsViewer(root)
    root.mainloop()