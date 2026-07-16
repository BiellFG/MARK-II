from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

# Rota principal do nosso site
@app.route('/')
def home():
    # Aqui é onde o Python faz o trabalho duro. 
    # Por enquanto, vamos mandar variáveis de teste para o HTML. 
    # No próximo passo, vamos plugar o seu Excel aqui!
    dados_resumo = {
        "mes": "Julho",
        "total_gasto": "R$ 2.363,50",
        "maior_despesa": "Aluguel (R$ 6.200,00)"
    }
    
    # O Python "renderiza" o HTML e injeta os dados nele
    return render_template('index.html', dados=dados_resumo)

if __name__ == '__main__':
    # O debug=True faz o site atualizar sozinho quando salvamos o código!
    app.run(debug=True)