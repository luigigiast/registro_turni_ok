from flask import Flask, render_template, request, redirect, session, send_file
import json
import csv
import os
from datetime import datetime, timedelta
import pandas as pd
import holidays

app = Flask(__name__)
app.secret_key = 'chiave_super_sicura'

# Carica utenti
with open('users.json', 'r') as f:
    utenti = json.load(f)

# Crea CSV se non esiste
if not os.path.exists('presenze.csv'):
    with open('presenze.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Nome', 'Data', 'Ora Ingresso', 'Ora Uscita', 'Durata'])

@app.route('/', methods=['GET', 'POST'])
def login():
    messaggio = None
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user_agent = request.headers.get('User-Agent')

        user = utenti.get(username)

        if user and user['password'] == password:
            # Se il dispositivo non è registrato, salvalo
            if 'device_id' not in user:
                utenti[username]['device_id'] = user_agent
                with open('users.json', 'w') as f:
                    json.dump(utenti, f, indent=4)

            # Se il dispositivo è diverso da quello registrato, blocca l’accesso
            elif user['device_id'] != user_agent:
                messaggio = "Accesso negato: questo account è già registrato su un altro dispositivo."
                return render_template('login.html', messaggio=messaggio)

            session['username'] = username
            session['nome'] = user['nome']
            session['ruolo'] = user['ruolo']
            session['ora_ingresso'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            messaggio = "Buon lavoro ed in bocca al lupo!"

            if user['ruolo'] == 'admin':
                return redirect('/admin')
            else:
                return render_template('login.html', messaggio=messaggio, fine_turno=True)
        else:
            messaggio = "Credenziali errate. Riprova."
    return render_template('login.html', messaggio=messaggio)

@app.route('/fine-turno', methods=['POST'])
def fine_turno():
    if 'username' in session and session.get('ruolo') == 'dipendente':
        ora_ingresso = datetime.strptime(session['ora_ingresso'], '%Y-%m-%d %H:%M:%S')
        ora_uscita = datetime.now()
        durata = ora_uscita - ora_ingresso

        if durata > timedelta(hours=8):
            durata = timedelta(hours=1)

        durata_str = str(durata).split('.')[0]

        with open('presenze.csv', 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([session['nome'], ora_ingresso.date(), ora_ingresso.time(), ora_uscita.time(), durata_str])

        messaggio = f"Spero sia andata bene, buona giornata a domani! Hai lavorato per: {durata_str}"
        session.clear()
        return f"<h2 style='text-align:center'>{messaggio}</h2><br><div style='text-align:center'><a href='/'>Torna al login</a></div>"

    return redirect('/')

@app.route('/admin')
def admin():
    if 'username' in session and session.get('ruolo') == 'admin':
        dati = []
        if os.path.exists('presenze.csv'):
            with open('presenze.csv', 'r') as f:
                reader = csv.reader(f)
                next(reader)  # salta intestazione
                for row in reader:
                    dati.append(row)

        return render_template('admin.html', dati=dati)
    return redirect('/')

@app.route('/download-csv')
def download_csv():
    if 'username' in session and session.get('ruolo') == 'admin':
        return send_file('presenze.csv', as_attachment=True)
    return redirect('/')

@app.route('/download-xlsx')
def download_xlsx():
    if 'username' in session and session.get('ruolo') == 'admin':
        df = pd.read_csv('presenze.csv')

        # Aggiunta colonna Giorno Festivo
        it_holidays = holidays.IT(years=datetime.now().year)
        df['Data'] = pd.to_datetime(df['Data'])
        df['Festivo o Domenica'] = df['Data'].apply(
            lambda x: 'SI' if x.weekday() == 6 or x in it_holidays else 'NO'
        )

        # Creazione del DataFrame con giorni del mese come colonne
        giorni_del_mese = [i for i in range(1, 32)]  # Giorni da 1 a 31
        giorni_colonne = {i: f'Giorno {i}' for i in giorni_del_mese}

        # Crea una tabella vuota con i giorni come colonne
        presenze_mese = pd.DataFrame(columns=giorni_colonne.values(), index=df['Nome'].unique())

        # Ordinamento alfabetico per cognome e nome
        df['Nome Completo'] = df['Nome'].apply(lambda x: ' '.join(x.split()[::-1]))  # Cognome Nome
        df = df.sort_values(by='Nome Completo')

        # Riempimento della tabella con le presenze
        for index, row in df.iterrows():
            giorno = row['Data'].day
            nome_completo = row['Nome Completo']
            durata = row['Durata']

            # Aggiungi la durata nella tabella del mese
            presenze_mese.at[nome_completo, f'Giorno {giorno}'] = durata

        # Aggiungi la colonna per festivo o domenica
        presenze_mese['Festivo o Domenica'] = df.groupby('Nome Completo')['Festivo o Domenica'].first()

        # Salva in Excel
        presenze_mese.to_excel('presenze_mese.xlsx', index=True)
        return send_file('presenze_mese.xlsx', as_attachment=True)
    return redirect('/')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)