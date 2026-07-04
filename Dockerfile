# syntax=docker/dockerfile:1

# --- Etape 1 : construction des dependances -------------------------------
FROM python:3.12-slim-bookworm AS builder

WORKDIR /app

# Environnement isole (venv) copie tel quel dans l'image finale : evite
# d'embarquer pip, les caches et les outils de build dans l'image finale.
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt gunicorn


# --- Etape 2 : image finale, minimale --------------------------------------
FROM python:3.12-slim-bookworm

# Mises a jour de securite du systeme de base uniquement (pas d'installation
# de paquets supplementaires : l'image reste minimale).
RUN apt-get update \
    && apt-get upgrade -y \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Utilisateur non-root dedie : l'application ne tourne jamais en root.
# --home-dir /app (au lieu de /home/appuser, jamais cree) : evite les
# erreurs de permission des outils qui ecrivent dans $HOME.
RUN groupadd --gid 1000 appuser \
    && useradd --uid 1000 --gid appuser --shell /usr/sbin/nologin --home-dir /app appuser \
    && chown appuser:appuser /app

COPY --from=builder /opt/venv /opt/venv

ENV PATH="/opt/venv/bin:$PATH" \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOME=/app \
    TEMPLATES_DIR=/app/modeles

COPY --chown=appuser:appuser . .

# Dossiers ecrits a l'execution (donnees generees, potentiellement montes en
# volume) : crees a l'avance avec les bons droits pour l'utilisateur non-root.
RUN mkdir -p web_runs releves_generes \
    && chown -R appuser:appuser web_runs releves_generes

USER appuser

EXPOSE 5000

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:5000/', timeout=2)" || exit 1

# Serveur WSGI de production (le serveur de dev Flask n'est pas utilise ici).
# --preload : app.py est importe une seule fois par le processus principal,
# donc une erreur au demarrage (ex. modele Excel manquant) arrete le
# conteneur immediatement avec un message clair, au lieu d'un redemarrage
# en boucle des workers.
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "--preload", "app:app"]
