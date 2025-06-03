import customtkinter as ctk
from tkinter import messagebox, ttk
import pandas as pd

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class OlympicMedalsViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("🏆 Visualizador de Medalhas Olímpicas")
        self.root.geometry("1100x750")
        
        # Estados e dados
        self.tabela_visivel = False
        self.df = self.carregar_dados()
        self.anos_disponiveis = sorted(self.df['year'].unique()) if not self.df.empty else []
        self.paises_disponiveis = sorted(self.df['country'].unique()) if not self.df.empty else []
        
        # Interface
        self.criar_interface()
    
    def carregar_dados(self):
        try:
            return pd.read_excel("world_olympedia_olympics_game_medal_tally.xlsx")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o Excel: {e}")
            return pd.DataFrame()
    
    def criar_interface(self):
        # Frame principal scrollável
        self.main_frame = ctk.CTkScrollableFrame(self.root, fg_color="#0d1117")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Título
        ctk.CTkLabel(self.main_frame, text="🏆 ANÁLISE DE MEDALHAS OLÍMPICAS", 
                    font=ctk.CTkFont(size=28, weight="bold"), text_color="#3b82f6").pack(pady=(0, 30))
        
        # Consultas
        consultas = [
            ("🥇 Top 10 países com mais medalhas em determinado ano", self.criar_consulta_ano),
            ("🌍 Total medalhas por país", self.criar_consulta_pais),
            ("🏆 Top 10 com mais mais medalhas de ouros", self.criar_consulta_ouros),
            ("🥈 Países sem medalhas de ouro", self.criar_consulta_sem_ouro),
            ("🏟 Edição mais competitiva", self.criar_consulta_competitiva),
            ("⭐ Países com medalhas em todas as edições", self.criar_consulta_todas_edicoes)
        ]
        
        for titulo, metodo in consultas:
            metodo(titulo)
        
        # Botão tabela completa
        self.botao_tabela = ctk.CTkButton(self.main_frame, text="📋 Mostrar Tabela Completa",
                                         command=self.alternar_tabela, width=200, height=45,
                                         fg_color="#1f538d", corner_radius=25)
        self.botao_tabela.pack(pady=30)
        
        # Frame da tabela
        self.criar_tabela()
    
    def criar_card(self, titulo):
        """Cria card padronizado para consultas"""
        card = ctk.CTkFrame(self.main_frame, fg_color="#161b22", border_color="#1f538d", 
                           border_width=2, corner_radius=15)
        card.pack(fill="x", pady=15, padx=10)
        
        ctk.CTkLabel(card, text=titulo, font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=20, pady=(20, 10))
        
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="x", padx=20, pady=(0, 20))
        return content
    
    def criar_botao_status(self, parent, texto_botao, comando, status_inicial, cor="#3b82f6"):
        """Cria botão e label de status padronizados"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=10)
        
        btn = ctk.CTkButton(frame, text=texto_botao, command=comando, fg_color=cor, 
                           hover_color="#2563eb", width=150, height=35)
        btn.pack(side="left", padx=10)
        
        status = ctk.CTkLabel(frame, text=status_inicial, text_color="#7d8590")
        status.pack(side="left", padx=20)
        return status
    
    def criar_consulta_ano(self, titulo):
        content = self.criar_card(titulo)
        frame = ctk.CTkFrame(content, fg_color="transparent")
        frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(frame, text="Ano:", text_color="#e6edf3").pack(side="left", padx=(0, 10))
        self.combo_ano = ctk.CTkComboBox(frame, values=[str(a) for a in self.anos_disponiveis], width=120)
        self.combo_ano.pack(side="left", padx=10)
        if self.anos_disponiveis:
            self.combo_ano.set(str(self.anos_disponiveis[0]))
        
        ctk.CTkButton(frame, text="🔍 Buscar", command=self.filtrar_por_ano, width=100).pack(side="left", padx=10)
        self.status_ano = ctk.CTkLabel(frame, text="Selecione um ano", text_color="#7d8590")
        self.status_ano.pack(side="left", padx=20)
    
    def criar_consulta_pais(self, titulo):
        content = self.criar_card(titulo)
        frame = ctk.CTkFrame(content, fg_color="transparent")
        frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(frame, text="País:", text_color="#e6edf3").pack(side="left", padx=(0, 10))
        self.combo_pais = ctk.CTkComboBox(frame, values=self.paises_disponiveis, width=200)
        self.combo_pais.pack(side="left", padx=10)
        if self.paises_disponiveis:
            self.combo_pais.set(self.paises_disponiveis[0])
        
        ctk.CTkButton(frame, text="🔍 Buscar", command=self.filtrar_por_pais, width=100).pack(side="left", padx=10)
        self.status_pais = ctk.CTkLabel(frame, text="Selecione um país", text_color="#7d8590")
        self.status_pais.pack(side="left", padx=20)
    
    def criar_consulta_ouros(self, titulo):
        content = self.criar_card(titulo)
        self.status_ouro = self.criar_botao_status(content, "🏆 Buscar", self.buscar_pais_mais_ouro,
                                                  "Clique para buscar", "#fbbf24")
    
    def criar_consulta_sem_ouro(self, titulo):
        content = self.criar_card(titulo)
        self.status_sem_ouro = self.criar_botao_status(content, "🔍 Sem Ouro", self.buscar_sem_ouro,
                                                      "Clique para buscar", "#dc2626")
    
    def criar_consulta_competitiva(self, titulo):
        content = self.criar_card(titulo)
        self.status_competitiva = self.criar_botao_status(content, "📈 Buscar", self.buscar_edicao_mais_competitiva,
                                                         "Clique para buscar", "#16a34a")
    
    def criar_consulta_todas_edicoes(self, titulo):
        content = self.criar_card(titulo)
        self.status_todas = self.criar_botao_status(content, "⭐ Buscar", self.buscar_paises_todas_edicoes,
                                                   "Clique para buscar", "#7c3aed")
    
    def criar_tabela(self):
        # Frame da tabela (inicialmente oculto)
        self.frame_tabela = ctk.CTkFrame(self.main_frame, fg_color="#161b22", border_color="#1f538d",
                                        border_width=2, corner_radius=15)
        
        # Título da tabela
        self.table_title = ctk.CTkLabel(self.frame_tabela, text="📊 RESULTADOS",
                                       font=ctk.CTkFont(size=18, weight="bold"), text_color="#3b82f6")
        
        # Frame para treeview
        tree_frame = ctk.CTkFrame(self.frame_tabela, fg_color="transparent")
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview", background="#161b22", foreground="#e6edf3",
                       fieldbackground="#161b22", borderwidth=0)
        style.configure("Custom.Treeview.Heading", background="#1f538d", foreground="white")
        style.map("Custom.Treeview", background=[('selected', '#3b82f6')])
        
        # Treeview com scrollbars
        self.colunas = list(self.df.columns) if not self.df.empty else []
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        self.tree = ttk.Treeview(tree_frame, columns=self.colunas, show='headings',
                                style="Custom.Treeview", yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Configurar colunas
        for col in self.colunas:
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=120, anchor='center')
        
        # Layout
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree_frame = tree_frame
    
    def atualizar_treeview(self, df_novo):
        # Limpar dados
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Atualizar colunas se necessário
        if list(df_novo.columns) != self.colunas:
            self.tree["columns"] = list(df_novo.columns)
            for col in df_novo.columns:
                self.tree.heading(col, text=col.capitalize())
                self.tree.column(col, width=120, anchor='center')
            self.colunas = list(df_novo.columns)
        
        # Inserir dados
        for _, row in df_novo.iterrows():
            self.tree.insert("", "end", values=list(row))
    
    def mostrar_tabela(self):
        if not self.tabela_visivel:
            self.table_title.pack(pady=(20, 10))
            self.tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
            self.frame_tabela.pack(fill="both", expand=True, pady=20, padx=10)
            self.botao_tabela.configure(text="🗂 Ocultar Tabela")
            self.tabela_visivel = True
    
    def ocultar_tabela(self):
        if self.tabela_visivel:
            self.frame_tabela.pack_forget()
            self.botao_tabela.configure(text="📋 Mostrar Tabela Completa")
            self.tabela_visivel = False
    
    def alternar_tabela(self):
        if self.tabela_visivel:
            self.ocultar_tabela()
        else:
            self.mostrar_tabela()
            self.atualizar_treeview(self.df)
    
    # Métodos de consulta otimizados
    def agrupar_medalhas(self, df, colunas_grupo):
        """Método genérico para agrupar medalhas"""
        return df.groupby(colunas_grupo).agg({
            'gold': 'sum', 'silver': 'sum', 'bronze': 'sum', 'total': 'sum'
        }).reset_index()
    
    def filtrar_por_ano(self):
        if not self.combo_ano.get():
            messagebox.showwarning("Atenção", "Selecione um ano.")
            return
        
        ano = int(self.combo_ano.get())
        df_filtrado = self.df[self.df['year'] == ano]
        
        if df_filtrado.empty:
            self.status_ano.configure(text="❌ Nenhum dado encontrado")
            return
        
        resumo = self.agrupar_medalhas(df_filtrado, ['country', 'country_noc'])
        top10 = resumo.sort_values(by='total', ascending=False).head(10)
        
        self.mostrar_tabela()
        self.atualizar_treeview(top10)
        self.status_ano.configure(text=f"✅ Top 10 países em {ano}")
    
    def filtrar_por_pais(self):
        if not self.combo_pais.get():
            messagebox.showwarning("Atenção", "Selecione um país.")
            return
        
        pais = self.combo_pais.get()
        df_filtrado = self.df[self.df['country'] == pais]
        
        if df_filtrado.empty:
            self.status_pais.configure(text="❌ Nenhum dado encontrado")
            return
        
        resumo = self.agrupar_medalhas(df_filtrado, ['year', 'country', 'country_noc'])
        resumo = resumo.sort_values(by='year', ascending=False)
        
        self.mostrar_tabela()
        self.atualizar_treeview(resumo)
        self.status_pais.configure(text=f"✅ Medalhas de {pais}")
    
    def buscar_pais_mais_ouro(self):
        if self.df.empty:
            self.status_ouro.configure(text="❌ Arquivo vazio")
            return
        
        resumo = self.agrupar_medalhas(self.df, ['country', 'country_noc'])
        top_ouro = resumo.sort_values(by='gold', ascending=False).head(10)
        
        self.mostrar_tabela()
        self.atualizar_treeview(top_ouro)
        self.status_ouro.configure(text="🏆 Top 10 países com mais ouros")
    
    def buscar_sem_ouro(self):
        if self.df.empty:
            self.status_sem_ouro.configure(text="❌ Arquivo vazio")
            return
        
        resumo = self.agrupar_medalhas(self.df, ['country', 'country_noc'])
        sem_ouro = resumo[resumo['gold'] == 0].sort_values(by='total', ascending=False)
        
        if sem_ouro.empty:
            self.status_sem_ouro.configure(text="ℹ Todos já ganharam ouro")
            return
        
        self.mostrar_tabela()
        self.atualizar_treeview(sem_ouro)
        self.status_sem_ouro.configure(text=f"📊 {len(sem_ouro)} Países sem medalhas de ouro")
    
    def buscar_edicao_mais_competitiva(self):
        if self.df.empty:
            self.status_competitiva.configure(text="❌ Arquivo vazio")
            return
        
        df_filtrado = self.df[self.df['total'] > 0]
        contagem = df_filtrado.groupby('year')['country'].nunique().reset_index(name='paises_com_medalha')
        ano_max = contagem.sort_values(by='paises_com_medalha', ascending=False).iloc[0]['year']
        
        paises_ano = df_filtrado[df_filtrado['year'] == ano_max]
        resumo = self.agrupar_medalhas(paises_ano, ['country', 'country_noc'])
        resumo = resumo.sort_values(by='total', ascending=False)
        
        self.mostrar_tabela()
        self.atualizar_treeview(resumo)
        self.status_competitiva.configure(text=f"🏟 {len(resumo)} países em {int(ano_max)}")
    
    def buscar_paises_todas_edicoes(self):
        if self.df.empty:
            self.status_todas.configure(text="❌ Arquivo vazio")
            return
        
        anos = set(self.df['year'].unique())
        df_filtrado = self.df[self.df['total'] > 0]
        pais_edicoes = df_filtrado.groupby('country')['year'].apply(set).reset_index()
        paises_todas = pais_edicoes[pais_edicoes['year'].apply(lambda x: anos.issubset(x))]
        paises_todas = paises_todas[['country']].sort_values(by='country')
        
        if paises_todas.empty:
            self.status_todas.configure(text="ℹ Nenhum país em todas edições")
            return
        
        self.mostrar_tabela()
        self.atualizar_treeview(paises_todas)
        self.status_todas.configure(text=f"⭐ {len(paises_todas)} Países com medalhas em todas as edições")

if __name__ == "__main__":
    from tkinter import messagebox
    import tkinter as tk
    tk.Tk().withdraw()
    messagebox.showwarning("Acesso negado", "Por favor, inicie o programa através do login.")