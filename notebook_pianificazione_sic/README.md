# 📋 Pianificazione Corsi 2026 - Notebook Standalone

Questa cartella contiene tutto il necessario per utilizzare il notebook Jupyter in modo autonomo.

## 📦 Contenuto

```
notebook_standalone/
├── pianificazione_completa.ipynb      # Notebook principale
├── crea_pianificazione_smart.py       # Modulo creazione Excel
├── genera_stampe_pdf.py               # Modulo generazione PDF
├── genera_stampe_pdf_filtrati.py      # Modulo PDF filtrati
├── requirements.txt                   # Dipendenze Python
└── README.md                          # Questo file
```

## 🚀 Come usare

### 1. Installazione dipendenze

```bash
pip install -r requirements.txt
```

Oppure:
```bash
pip install openpyxl reportlab
```

### 2. Avvio Jupyter

```bash
jupyter notebook pianificazione_completa.ipynb
```

Oppure usa JupyterLab:
```bash
jupyter lab
```

### 3. Utilizzo del Notebook

1. Esegui le celle in ordine dall'alto verso il basso
2. **Sezione 5**: Crea il file Excel di pianificazione
3. **Apri il file Excel** e compila i dati
4. **Sezione 6**: Genera tutti i PDF (o usa le sezioni specifiche)
5. **Sezione 7**: Gestisci i file generati

## 📁 File Generati

Dopo l'esecuzione, nella cartella verranno creati:

- `Pianificazione_Corsi_2026.xlsx` - File Excel con i dati
- `stampe_pdf/` - Cartella con tutti i PDF generati
  - `Prenotazione_Aule_2026.pdf`
  - `Programma_Formatore_*.pdf`
  - `Programma_Corso_*.pdf`
  - `Piano_Settimanale_*.pdf`

## ✅ Vantaggi

- ✅ **Autonomo**: Non richiede server web o altre applicazioni
- ✅ **Portabile**: Funziona su Windows, macOS e Linux
- ✅ **Completo**: Tutte le funzionalità dell'app web
- ✅ **Interattivo**: Esegui celle singolarmente o tutto insieme

## 🔧 Requisiti

- Python 3.8+
- Jupyter Notebook o JupyterLab
- openpyxl
- reportlab

## 📝 Note

Questa versione è completamente indipendente dall'applicazione web principale.
Puoi spostare questa cartella ovunque e continuerà a funzionare.

---

**🎯 Pronto all'uso!**
