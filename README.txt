MedSafe AI — Render + GitHub Ready

DEPLOY
1. Push this folder to a GitHub repository.
2. In Render, choose New -> Blueprint and select the repository.
3. Render reads render.yaml and creates the web service.
4. The blueprint generates MEDSAFE_SECRET_KEY automatically.
5. The bundled SQLite seed database is copied to /var/data on first start.

LOGIN
Username: admin
Password: admin123

Hospitals can register separate accounts from the registration page.

IMPORTANT PRODUCTION NOTE
This application handles patient and medication data. The included admin password is for demonstration/testing and MUST be changed before real-world use. For a public production deployment, also configure stronger operational security, HTTPS (Render provides this), backups, access controls, audit logging, and a managed database if required by your deployment.

LOCAL RUN
py -m pip install -r requirements.txt
py app.py

Production command:
gunicorn --workers 2 --threads 4 --timeout 120 app:app
