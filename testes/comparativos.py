import matplotlib.pyplot as plt
import numpy as np

# --- DADOS BASE ---
peso_real = 41550
pesos_totais = {
    "Remessa Real": peso_real,
    "Simulador Antigo": 42120,
    "Simulador Novo": 41255
}
pesos_produtos = {
    "Remessa Real": 23930,
    "Simulador Antigo": 24492,
    "Simulador Novo": 23635.17
}
margem_antigo = (41690.88, 43796.48)
margem_novo = (40842.62, 42080.28)

# --- 1. GRÁFICO DE PESO TOTAL ---
fig, ax = plt.subplots(figsize=(10, 6))
labels = list(pesos_totais.keys())
valores = list(pesos_totais.values())
bars = ax.bar(labels, valores, color=['steelblue', 'orange', 'green'])

for bar in bars:
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, yval - 800, f"{yval:.0f}", ha='center', va='top', color='white', fontsize=10)

ax.axhline(margem_antigo[0], color='orange', linestyle='--', label='Margem Antigo (mín)')
ax.axhline(margem_antigo[1], color='orange', linestyle='--', label='Margem Antigo (máx)')
ax.axhline(margem_novo[0], color='green', linestyle=':', label='Margem Novo (mín)')
ax.axhline(margem_novo[1], color='green', linestyle=':', label='Margem Novo (máx)')

ax.set_title("Peso Total da Remessa - Comparativo")
ax.set_ylabel("Peso (kg)")
ax.legend()
plt.tight_layout()
plt.show()

# --- 2. GRÁFICO DE PESO DOS PRODUTOS ---
fig, ax2 = plt.subplots(figsize=(10, 5))
labels_prod = list(pesos_produtos.keys())
valores_prod = list(pesos_produtos.values())
bars_prod = ax2.bar(labels_prod, valores_prod, color=['lightblue', 'orange', 'lightgreen'])

for bar in bars_prod:
    yval = bar.get_height()
    ax2.text(bar.get_x() + bar.get_width()/2, yval - 500, f"{yval:.0f}", ha='center', va='top', color='black', fontsize=10)

ax2.set_title("Peso Só dos Produtos - Comparativo")
ax2.set_ylabel("Peso (kg)")
plt.tight_layout()
plt.show()

# --- 3. GRÁFICO GAUGE DAS DIFERENÇAS EM % ---
dif_antigo = ((pesos_totais["Simulador Antigo"] - peso_real) / peso_real) * 100
dif_novo = ((pesos_totais["Simulador Novo"] - peso_real) / peso_real) * 100

def plot_gauge(ax, value, title, color):
    ax.set_xlim(-1.5, 1.5)
    ax.set_ylim(-1.5, 1.5)
    ax.axis('off')
    ax.set_title(title, fontsize=12)
    fundo = plt.Circle((0, 0), 1.0, color='lightgray', zorder=0)
    ax.add_patch(fundo)

    for angle in np.linspace(-90, 90, 100):
        x = np.cos(np.radians(angle))
        y = np.sin(np.radians(angle))
        if -2 <= angle / 90 * 100 <= 2:
            ax.plot([0, x], [0, y], color='green', linewidth=3, alpha=0.6)
        else:
            ax.plot([0, x], [0, y], color='gray', linewidth=1, alpha=0.3)

    angle = value / 100 * 90
    angle = max(-90, min(90, angle))
    x = np.cos(np.radians(angle))
    y = np.sin(np.radians(angle))
    ax.arrow(0, 0, x, y, width=0.05, head_width=0.1, head_length=0.1, fc=color, ec=color)
    ax.text(0, -1.3, f"{value:.2f}%", ha='center', fontsize=12, color=color)

fig, axs = plt.subplots(1, 2, figsize=(10, 5), subplot_kw={'aspect': 'equal'})
plot_gauge(axs[0], dif_antigo, "Simulador Antigo", "orange")
plot_gauge(axs[1], dif_novo, "Simulador Novo", "green")
plt.suptitle("Diferença Percentual em Relação ao Peso Real", fontsize=14)
plt.tight_layout()
plt.show()

# --- 4. FAIXA DE TOLERÂNCIA EM % E KG ---
antigo_min_pct = ((margem_antigo[0] - peso_real) / peso_real) * 100
antigo_max_pct = ((margem_antigo[1] - peso_real) / peso_real) * 100
novo_min_pct = ((margem_novo[0] - peso_real) / peso_real) * 100
novo_max_pct = ((margem_novo[1] - peso_real) / peso_real) * 100

antigo_min_kg = margem_antigo[0] - peso_real
antigo_max_kg = margem_antigo[1] - peso_real
novo_min_kg = margem_novo[0] - peso_real
novo_max_kg = margem_novo[1] - peso_real

labels = ['Simulador Antigo', 'Simulador Novo']
minimos_pct = [antigo_min_pct, novo_min_pct]
maximos_pct = [antigo_max_pct, novo_max_pct]
minimos_kg = [antigo_min_kg, novo_min_kg]
maximos_kg = [antigo_max_kg, novo_max_kg]

x = np.arange(len(labels))
largura = 0.35

fig, ax3 = plt.subplots(figsize=(10, 6))
bars1 = ax3.bar(x - largura/2, minimos_pct, largura, label='Tolerância Mínima (%)', color='skyblue')
bars2 = ax3.bar(x + largura/2, maximos_pct, largura, label='Tolerância Máxima (%)', color='orange')

for i, bar in enumerate(bars1):
    pct = minimos_pct[i]
    kg = minimos_kg[i]
    ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height()/2,
             f"{pct:.2f}%\n({kg:.0f} kg)", ha='center', va='center', fontsize=10, color='black')

for i, bar in enumerate(bars2):
    pct = maximos_pct[i]
    kg = maximos_kg[i]
    ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height()/2,
             f"{pct:.2f}%\n({kg:.0f} kg)", ha='center', va='center', fontsize=10, color='black')

ax3.set_ylabel("Variação em relação ao peso real (%)")
ax3.set_title("Faixa de Aceitação dos Simuladores em % e KG")
ax3.set_xticks(x)
ax3.set_xticklabels(labels)
ax3.axhline(0, color='gray', linestyle='--')
ax3.legend()
plt.tight_layout()
plt.show()
