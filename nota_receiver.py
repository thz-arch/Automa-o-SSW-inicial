from flask import Flask, request, jsonify
import subprocess
import sys
import os

app = Flask(__name__)

@app.route('/nota', methods=['POST'])
def receber_nota():
    data = request.get_json()
    if not data or 'nota' not in data:
        return jsonify({'error': 'Campo "nota" não encontrado no JSON.'}), 400
    nota = str(data['nota']).strip()
    with open('nota_atual.txt', 'w', encoding='utf-8') as f:
        f.write(nota)
    # Dispara o script ssw_login.py em segundo plano e registra saída/erro
    python_exe = sys.executable  # Caminho absoluto do Python em uso
    script_path = os.path.join(os.path.dirname(__file__), 'ssw_login.py')
    with open('ssw_login_exec.log', 'a', encoding='utf-8') as log:
        log.write(f'\nTeste de escrita no log para nota {nota}\n')
        log.flush()
        try:
            log.write(f'Disparando {python_exe} {script_path} para nota {nota}...\n')
            log.flush()
            subprocess.Popen([python_exe, script_path], stdout=log, stderr=log)
        except Exception as e:
            log.write(f'Erro ao tentar disparar o script: {e}\n')
            log.flush()
    return jsonify({'message': f'Nota {nota} recebida e salva com sucesso!'}), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
