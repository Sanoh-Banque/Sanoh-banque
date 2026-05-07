from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
import sqlite3
import datetime
import os

app = Flask(__name__)
app.secret_key = 'sanoh-banque-secret-2026'

DB_PATH = 'banquepy.db'

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS comptes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero TEXT UNIQUE NOT NULL,
            titulaire TEXT NOT NULL,
            solde REAL DEFAULT 0.0,
            date_creation TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            compte_numero TEXT NOT NULL,
            titulaire TEXT NOT NULL,
            type TEXT NOT NULL,
            montant REAL NOT NULL,
            date TEXT NOT NULL,
            description TEXT
        );
    ''')
    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    conn = get_db()
    comptes = conn.execute('SELECT * FROM comptes').fetchall()
    operations = conn.execute('SELECT * FROM operations ORDER BY id DESC LIMIT 10').fetchall()
    total_fonds = conn.execute('SELECT SUM(solde) as total FROM comptes').fetchone()['total'] or 0
    ops_today = conn.execute("SELECT COUNT(*) as count FROM operations WHERE date LIKE ?", (datetime.date.today().isoformat() + '%',)).fetchone()['count']
    conn.close()
    return render_template('index.html', comptes=comptes, operations=operations, total_fonds=total_fonds, ops_today=ops_today)

@app.route('/nouveau-compte', methods=['GET', 'POST'])
def nouveau_compte():
    if request.method == 'POST':
        titulaire = request.form['titulaire']
        solde_initial = float(request.form.get('solde_initial', 0))
        numero = 'SB' + datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        date_creation = datetime.datetime.now().isoformat()
        conn = get_db()
        try:
            conn.execute('INSERT INTO comptes (numero, titulaire, solde, date_creation) VALUES (?, ?, ?, ?)',
                        (numero, titulaire, solde_initial, date_creation))
            if solde_initial > 0:
                conn.execute('INSERT INTO operations (compte_numero, titulaire, type, montant, date) VALUES (?, ?, ?, ?, ?)',
                            (numero, titulaire, 'Dépôt initial', solde_initial, date_creation))
            conn.commit()
            flash(f'Compte {numero} créé avec succès !', 'success')
        except Exception as e:
            flash(f'Erreur : {str(e)}', 'error')
        finally:
            conn.close()
        return redirect(url_for('index'))
    return render_template('nouveau_compte.html')

@app.route('/depot', methods=['GET', 'POST'])
def depot():
    conn = get_db()
    comptes = conn.execute('SELECT * FROM comptes').fetchall()
    if request.method == 'POST':
        numero = request.form['numero']
        montant = float(request.form['montant'])
        compte = conn.execute('SELECT * FROM comptes WHERE numero = ?', (numero,)).fetchone()
        if compte and montant > 0:
            conn.execute('UPDATE comptes SET solde = solde + ? WHERE numero = ?', (montant, numero))
            conn.execute('INSERT INTO operations (compte_numero, titulaire, type, montant, date) VALUES (?, ?, ?, ?, ?)',
                        (numero, compte['titulaire'], 'Dépôt', montant, datetime.datetime.now().isoformat()))
            conn.commit()
            flash(f'Dépôt de {montant:.2f} DA effectué !', 'success')
        else:
            flash('Compte introuvable ou montant invalide.', 'error')
        conn.close()
        return redirect(url_for('index'))
    conn.close()
    return render_template('depot.html', comptes=comptes)

@app.route('/retrait', methods=['GET', 'POST'])
def retrait():
    conn = get_db()
    comptes = conn.execute('SELECT * FROM comptes').fetchall()
    if request.method == 'POST':
        numero = request.form['numero']
        montant = float(request.form['montant'])
        compte = conn.execute('SELECT * FROM comptes WHERE numero = ?', (numero,)).fetchone()
        if compte and montant > 0 and compte['solde'] >= montant:
            conn.execute('UPDATE comptes SET solde = solde - ? WHERE numero = ?', (montant, numero))
            conn.execute('INSERT INTO operations (compte_numero, titulaire, type, montant, date) VALUES (?, ?, ?, ?, ?)',
                        (numero, compte['titulaire'], 'Retrait', montant, datetime.datetime.now().isoformat()))
            conn.commit()
            flash(f'Retrait de {montant:.2f} DA effectué !', 'success')
        elif compte and compte['solde'] < montant:
            flash('Solde insuffisant !', 'error')
        else:
            flash('Compte introuvable ou montant invalide.', 'error')
        conn.close()
        return redirect(url_for('index'))
    conn.close()
    return render_template('retrait.html', comptes=comptes)

@app.route('/virement', methods=['GET', 'POST'])
def virement():
    conn = get_db()
    comptes = conn.execute('SELECT * FROM comptes').fetchall()
    if request.method == 'POST':
        source = request.form['source']
        destination = request.form['destination']
        montant = float(request.form['montant'])
        c_source = conn.execute('SELECT * FROM comptes WHERE numero = ?', (source,)).fetchone()
        c_dest = conn.execute('SELECT * FROM comptes WHERE numero = ?', (destination,)).fetchone()
        if c_source and c_dest and montant > 0 and c_source['solde'] >= montant and source != destination:
            now = datetime.datetime.now().isoformat()
            conn.execute('UPDATE comptes SET solde = solde - ? WHERE numero = ?', (montant, source))
            conn.execute('UPDATE comptes SET solde = solde + ? WHERE numero = ?', (montant, destination))
            conn.execute('INSERT INTO operations (compte_numero, titulaire, type, montant, date, description) VALUES (?, ?, ?, ?, ?, ?)',
                        (source, c_source['titulaire'], 'Virement envoyé', montant, now, f'Vers {destination}'))
            conn.execute('INSERT INTO operations (compte_numero, titulaire, type, montant, date, description) VALUES (?, ?, ?, ?, ?, ?)',
                        (destination, c_dest['titulaire'], 'Virement reçu', montant, now, f'De {source}'))
            conn.commit()
            flash(f'Virement de {montant:.2f} DA effectué !', 'success')
        else:
            flash('Erreur : vérifiez les comptes, le montant et le solde.', 'error')
        conn.close()
        return redirect(url_for('index'))
    conn.close()
    return render_template('virement.html', comptes=comptes)

@app.route('/historique')
def historique():
    conn = get_db()
    operations = conn.execute('SELECT * FROM operations ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('historique.html', operations=operations)

@app.route('/tous-les-comptes')
def tous_les_comptes():
    conn = get_db()
    comptes = conn.execute('SELECT * FROM comptes ORDER BY id DESC').fetchall()
    conn.close()
    return render_template('tous_les_comptes.html', comptes=comptes)

@app.route('/solde/<numero>')
def consulter_solde(numero):
    conn = get_db()
    compte = conn.execute('SELECT * FROM comptes WHERE numero = ?', (numero,)).fetchone()
    conn.close()
    return render_template('solde.html', compte=compte)

if __name__ == '__main__':
    app.run(debug=True)
    