import pandas as pd
import matplotlib.pyplot as plt
import sys
import os
from pathlib import Path

def verificar_dependencias():
    """Verifica se todas as dependências estão instaladas"""
    try:
        import pandas
        import matplotlib
        return True
    except ImportError:
        print("Erro: Dependências não instaladas. Execute:", file=sys.stderr)
        print("pip install pandas matplotlib", file=sys.stderr)
        return False

def carregar_dados(csv_path):
    """Carrega e valida os dados do CSV"""
    try:
        df = pd.read_csv(csv_path)
        
        # Verificar colunas necessárias
        required_cols = {'nome', 'fez_prova', 'nota'}
        if not required_cols.issubset(df.columns):
            print(f"Erro: CSV deve conter as colunas: {required_cols}", file=sys.stderr)
            return None
        
        # Converter fez_prova para booleano
        df['fez_prova'] = df['fez_prova'].astype(str).str.lower().map({'true': True, 'false': False, '1': True, '0': False})
        
        # Validar notas
        df['nota'] = pd.to_numeric(df['nota'], errors='coerce')
        df['nota'] = df['nota'].clip(0, 10)  # Garantir que notas estão entre 0 e 10
        
        return df
    
    except Exception as e:
        print(f"Erro ao carregar dados: {str(e)}", file=sys.stderr)
        return None

def gerar_graficos(csv_path, nome_sala, output_path):
    """Gera os gráficos e salva no arquivo de saída"""
    if not verificar_dependencias():
        return False
    
    # Verificar se arquivo de entrada existe
    if not Path(csv_path).exists():
        print(f"Erro: Arquivo CSV não encontrado: {csv_path}", file=sys.stderr)
        return False
    
    # Carregar dados
    df = carregar_dados(csv_path)
    if df is None:
        return False
    
    # Filtrar alunos que fizeram a prova
    df_provas = df[df['fez_prova'] == True]
    
    # Cálculos estatísticos com tratamento para divisão por zero
    total_alunos = len(df)
    fizeram_prova = len(df_provas)
    nao_fizeram = total_alunos - fizeram_prova
    
    try:
        media_sala = df_provas['nota'].mean() if fizeram_prova > 0 else 0
        acima_media = len(df_provas[df_provas['nota'] >= 6]) if fizeram_prova > 0 else 0
        abaixo_media = fizeram_prova - acima_media if fizeram_prova > 0 else 0
        
        # Criar figura com layout organizado
        plt.figure(figsize=(15, 10))
        plt.suptitle(f'Análise da Sala: {nome_sala}\nTotal de Alunos: {total_alunos}', 
                    fontsize=16, y=0.98)
        
        # Gráfico 1: Participação na prova
        ax1 = plt.subplot(2, 2, 1)
        if total_alunos > 0:
            participacao = [fizeram_prova, nao_fizeram]
            labels = [f'Fizeram ({fizeram_prova})', f'Não Fizeram ({nao_fizeram})']
            colors = ['#4CAF50', '#F44336']
            
            ax1.pie(participacao, labels=labels, colors=colors,
                   autopct=lambda p: f'{p:.1f}%\n({int(p/100*sum(participacao))})',
                   textprops={'fontsize': 10})
        else:
            ax1.text(0.5, 0.5, 'Nenhum aluno cadastrado', 
                    ha='center', va='center', fontsize=12)
        ax1.set_title('Participação na Prova', pad=20)
        
        # Gráfico 2: Desempenho
        ax2 = plt.subplot(2, 2, 2)
        if fizeram_prova > 0:
            desempenho = [acima_media, abaixo_media]
            labels = [f'Acima da média\n({acima_media})', f'Abaixo\n({abaixo_media})']
            colors = ['#2196F3', '#FF9800']
            
            ax2.pie(desempenho, labels=labels, colors=colors,
                   autopct=lambda p: f'{p:.1f}%',
                   textprops={'fontsize': 10})
        else:
            ax2.text(0.5, 0.5, 'Nenhum aluno\nfez a prova', 
                    ha='center', va='center', fontsize=12)
        ax2.set_title('Desempenho (média = 6.0)', pad=20)
        
        # Gráfico 3: Distribuição de notas
        ax3 = plt.subplot(2, 2, 3)
        if fizeram_prova > 0:
            n, bins, patches = ax3.hist(df_provas['nota'], bins=10, range=(0, 10), 
                                   color='#9C27B0', edgecolor='black')
            
            # Adicionar contagem em cada barra
            for i in range(len(patches)):
                ax3.text(bins[i] + 0.5, n[i] + 0.5, str(int(n[i])), 
                        ha='center', va='bottom')
            
            ax3.axvline(6, color='green', linestyle='-', linewidth=1, 
                       label='Mínimo (6.0)')
            ax3.axvline(media_sala, color='red', linestyle='--', linewidth=2, 
                       label=f'Média ({media_sala:.1f})')
            ax3.set_xlabel('Notas')
            ax3.set_ylabel('Quantidade de Alunos')
            ax3.legend()
            ax3.set_xlim(0, 10)
            ax3.grid(True, linestyle='--', alpha=0.5)
        else:
            ax3.text(0.5, 0.5, 'Nenhum aluno\nfez a prova', 
                    ha='center', va='center', fontsize=12)
        ax3.set_title('Distribuição de Notas', pad=20)
        
        # Gráfico 4: Estatísticas detalhadas
        ax4 = plt.subplot(2, 2, 4)
        stats_text = []
        stats_text.append(f"Total de Alunos: {total_alunos}")
        
        if total_alunos > 0:
            stats_text.append(f"Fizeram Prova: {fizeram_prova} ({fizeram_prova/total_alunos*100:.1f}%)")
            stats_text.append(f"Não Fizeram: {nao_fizeram} ({nao_fizeram/total_alunos*100:.1f}%)")
            
            if fizeram_prova > 0:
                stats_text.append(f"\nMédia da Sala: {media_sala:.1f}")
                stats_text.append(f"Maior Nota: {df_provas['nota'].max():.1f}")
                stats_text.append(f"Menor Nota: {df_provas['nota'].min():.1f}")
                stats_text.append(f"\nAcima/igual a 6: {acima_media} ({acima_media/fizeram_prova*100:.1f}%)")
                stats_text.append(f"Abaixo de 6: {abaixo_media} ({abaixo_media/fizeram_prova*100:.1f}%)")
        
        ax4.text(0.05, 0.95, "\n".join(stats_text), 
                fontsize=12, family='monospace',
                va='top', linespacing=1.5)
        ax4.set_title('Resumo Estatístico', pad=20)
        ax4.axis('off')
        
        # Ajustar layout e salvar
        plt.tight_layout(rect=[0, 0, 1, 0.95])
        plt.savefig(output_path, dpi=120, bbox_inches='tight')
        plt.close()
        
        print(f"Gráficos gerados com sucesso em: {output_path}")
        return True
        
    except Exception as e:
        print(f"Erro inesperado ao gerar gráficos: {str(e)}", file=sys.stderr)
        return False

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Uso correto: python gerar_graficos.py <arquivo_csv> <nome_sala> <arquivo_saida>", 
              file=sys.stderr)
        sys.exit(1)
    
    sucesso = gerar_graficos(sys.argv[1], sys.argv[2], sys.argv[3])
    sys.exit(0 if sucesso else 1)