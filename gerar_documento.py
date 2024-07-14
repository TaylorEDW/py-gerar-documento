from flask import Flask, render_template, request, send_file
import openpyxl
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preencher', methods=['POST'])
def preencher():
    nome = request.form['nome']
    demanda = request.form['demanda']
    modelo = request.form['modelo']
    inversor = request.form['inversor']
    inversores = request.form['inversores']
    placa = request.form['placa']
    potenciap = request.form['potenciap']
    quantidade = request.form['quantidade']
    proposta = request.form['proposta']

    # Carregar o arquivo
    book = openpyxl.load_workbook('MODELO ORÇAMENTO.xlsx')
    inicial_page = book['INICIAL']

    # Adicionar os dados nas células
    inicial_page['B7'] = nome
    inicial_page['G26'] = demanda
    inicial_page['E37'] = modelo
    inicial_page['E38'] = inversor
    inicial_page['E39'] = inversores
    inicial_page['H37'] = placa
    inicial_page['H38'] = potenciap
    inicial_page['H39'] = quantidade
    inicial_page['E59'] = proposta

    # Salvar alterações em memória
    output = BytesIO()
    book.save(output)
    output.seek(0)

    return send_file(output, download_name='MODELO ORÇAMENTO v2.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
