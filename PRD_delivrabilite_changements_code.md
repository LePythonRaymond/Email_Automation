# PRD — Durcissement délivrabilité de l'app d'envoi d'emails

**Produit :** MERCI RAYMOND · Raymongraphe (Streamlit + Gmail SMTP)
**Fichier principal concerné :** `email_automation_app.py` (~2210 lignes)
**Destinataire de ce doc :** Claude Code
**Auteur :** Taddeo
**Date :** 2026-05

---

## 1. Contexte

L'application envoie des emails en masse depuis 6 boîtes `@merciraymond.fr` via Gmail SMTP (`smtp.gmail.com:587`). L'audit de délivrabilité a identifié plusieurs manques qui peuvent dégrader la réputation du domaine. Ce PRD couvre les correctifs **côté code** à fort impact et faible effort.

**Hors périmètre de ce PRD** (gérés ailleurs) : la configuration DNS (SPF/DKIM/DMARC) et la création d'un sous-domaine dédié. La migration des mots de passe vers `st.secrets` est incluse ici car c'est un changement de code (le contenu des secrets sera fourni par Taddeo).

**Point d'attention critique pour l'implémentation :** il existe **DEUX boucles d'envoi quasi identiques** dans `main()` :
1. Envoi des emails « corrigés » (anciennement invalides) — autour des lignes **1887–1984**.
2. Envoi des emails « prêts » (remaining_emails) — autour des lignes **2060–2158**.

**Tous les changements ci-dessous doivent être appliqués aux DEUX boucles.** Pour éviter la duplication, l'implémentation privilégiée est de **factoriser l'envoi d'un email dans une fonction unique** réutilisée par les deux boucles (voir §6).

---

## 2. Objectifs

| Objectif | Mesure de succès |
|---|---|
| Conformité légale (RGPD/LCEN) + Gmail/Yahoo | Chaque email sortant porte un en-tête `List-Unsubscribe` + `List-Unsubscribe-Post` et un lien de désinscription visible. |
| Ne jamais réécrire à un désinscrit / bounce | Une liste de suppression persistante est vérifiée avant chaque envoi ; aucun email vers une adresse suppressed. |
| Réduire le pattern robotique | Intervalle d'envoi aléatoire entre 1 et 10 s. |
| Éviter les pics de volume / dépassement de quota | Plafond quotidien configurable par expéditeur, persistant sur la journée. |
| Sécuriser les identifiants | Plus aucun mot de passe en clair dans le code source. |

---

## 3. Exigences fonctionnelles

### F1 — En-têtes de désinscription (List-Unsubscribe)

À l'assemblage de chaque `msg_root` (actuellement lignes ~1894-1897 et ~2067-2070), ajouter :

```python
msg_root['List-Unsubscribe'] = (
    f'<mailto:desinscription@merciraymond.fr?subject=unsubscribe%20{email_data["email"]}>'
)
msg_root['List-Unsubscribe-Post'] = 'List-Unsubscribe=One-Click'
msg_root['Reply-To'] = sender_email
```

- **Version `mailto:` obligatoire** (fonctionne immédiatement, aucune infra requise).
- **Version `https://` (one-click) optionnelle / future** : à n'ajouter QUE si un endpoint de désinscription HTTP existe. Tant qu'il n'existe pas, ne PAS mettre d'URL https morte (un lien 404 nuit à la réputation). Laisser un TODO commenté.
- `Reply-To` = adresse de l'expéditeur (permet les réponses, signal positif).

**Critère d'acceptation F1 :** un email envoyé contient bien les 3 en-têtes ; la version texte et HTML sont inchangées par ailleurs.

### F2 — Lien de désinscription visible dans le corps

Ajouter automatiquement un **footer de désinscription** à TOUS les emails, indépendamment du contenu rédigé par l'utilisateur (le contenu utilisateur ne contient pas toujours de lien).

- Injecter le footer dans `self.html_template` (classe `EmailAutomation`, ~ligne 234), après `{logo_section}`, dans un `<div>` discret (petit, gris) :

```html
<div style="margin-top:24px; padding-top:12px; border-top:1px solid #eee; font-size:11px; color:#999;">
  Vous recevez cet email de MERCI RAYMOND.
  Pour ne plus en recevoir, <a href="mailto:desinscription@merciraymond.fr?subject=unsubscribe" style="color:#999;">cliquez ici pour vous désinscrire</a>.
</div>
```

- Le même texte doit apparaître dans la version texte (`html2text` le convertira automatiquement puisqu'il est dans le HTML — vérifier que c'est bien le cas).

**Critère d'acceptation F2 :** tout email rendu (HTML + texte) contient un lien/mention de désinscription, même si l'utilisateur n'en a pas mis dans son contenu.

### F3 — Liste de suppression persistante

Créer un module de gestion (idéalement un petit fichier `suppression.py`, ou des fonctions dans le fichier principal) :

```python
import json, os
from datetime import datetime

SUPPRESSION_FILE = "suppression_list.json"  # même dossier que le script

def _load_suppression() -> dict:
    if os.path.exists(SUPPRESSION_FILE):
        try:
            return json.load(open(SUPPRESSION_FILE, encoding="utf-8"))
        except Exception:
            return {}
    return {}

def is_suppressed(email: str) -> bool:
    return email.strip().lower() in _load_suppression()

def add_suppression(email: str, reason: str = "manual_unsubscribe"):
    data = _load_suppression()
    data[email.strip().lower()] = {"date": datetime.now().isoformat(), "reason": reason}
    json.dump(data, open(SUPPRESSION_FILE, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
```

Intégration :
- **Avant chaque envoi** dans les deux boucles : `if is_suppressed(email_data['email']): skipped_count += 1; continue` (NE PAS compter comme un échec).
- Filtrer aussi en amont, dès la construction de la liste des destinataires (`get_valid_emails_from_df`, ~ligne 412, et la liste `remaining_emails`) pour que les compteurs affichés soient justes.
- **Petite UI Streamlit** dans la sidebar ou un expander « Gestion désinscriptions » : un `text_area` pour coller des emails + bouton « Ajouter à la liste de suppression » qui appelle `add_suppression(..., reason="manual_unsubscribe")`. Permet de traiter manuellement les réponses reçues sur `desinscription@`.

**Critère d'acceptation F3 :** un email présent dans `suppression_list.json` n'est jamais envoyé (vérifié dans les deux boucles) ; l'UI permet d'ajouter des adresses ; le compteur « X destinataires ignorés (désinscrits) » est affiché.

### F4 — Intervalle d'envoi aléatoire (1 à 10 s)

Remplacer les deux `time.sleep(1)` (lignes ~1982 et ~2156) par :

```python
import random
if i < len(LISTE) - 1:        # garder la condition existante "pas de délai après le dernier"
    time.sleep(random.uniform(MIN_DELAY_S, MAX_DELAY_S))
```

- Conserver la logique « pas de délai après le dernier email ».
- Mettre les bornes dans des constantes en haut du fichier : `MIN_DELAY_S = 1`, `MAX_DELAY_S = 10`.

**Critère d'acceptation F4 :** l'intervalle entre deux envois est tiré aléatoirement dans [1, 10] s ; plus aucun `sleep(1)` fixe pour l'envoi.

### F5 — Plafond quotidien par expéditeur

Limiter le nombre d'envois par boîte et par jour, persistant sur la journée (même si on relance l'app).

```python
DAILY_CAP = 100  # configurable via l'UI

def _daily_count_file():
    return "daily_send_count.json"

def get_today_count(sender_email: str) -> int:
    from datetime import date
    today = date.today().isoformat()
    data = json.load(open(_daily_count_file(), encoding="utf-8")) if os.path.exists(_daily_count_file()) else {}
    return data.get(sender_email, {}).get(today, 0)

def incr_today_count(sender_email: str, n: int = 1):
    from datetime import date
    today = date.today().isoformat()
    data = json.load(open(_daily_count_file(), encoding="utf-8")) if os.path.exists(_daily_count_file()) else {}
    data.setdefault(sender_email, {})
    data[sender_email][today] = data[sender_email].get(today, 0) + n
    json.dump(data, open(_daily_count_file(), "w", encoding="utf-8"), indent=2)
```

Intégration dans les deux boucles :
- Avant d'envoyer : `if get_today_count(sender_email) >= DAILY_CAP: st.warning(...); break`.
- Après un envoi réussi : `incr_today_count(sender_email)`.
- Afficher dans l'UI le quota restant du jour pour l'expéditeur sélectionné.
- Rendre `DAILY_CAP` configurable via un `st.number_input` (valeur par défaut 100, max raisonnable 200).

**Critère d'acceptation F5 :** l'envoi s'arrête proprement quand le plafond du jour est atteint pour l'expéditeur ; le compteur survit à un redémarrage de l'app le même jour ; un nouveau jour remet le compteur à zéro.

### F6 — Migration des identifiants vers `st.secrets`

Les mots de passe Gmail (App Passwords) sont actuellement en clair dans `get_users()` / la liste `base` (~lignes 824-831). **À sortir du code.**

- Lire les credentials depuis `st.secrets` (Taddeo fournira `.streamlit/secrets.toml`).
- Structure proposée pour `secrets.toml` :

```toml
[users.salome]
name = "Salomé Cremona"
email = "salome.cremona@merciraymond.fr"
password = "xxxx xxxx xxxx xxxx"

[users.taddeo]
name = "Taddeo Carpinelli"
email = "taddeo.carpinelli@merciraymond.fr"
password = "xxxx xxxx xxxx xxxx"
# ... etc pour les 6 utilisateurs
```

- `get_users()` doit lire `st.secrets["users"]` et construire la liste. **Supprimer toutes les valeurs `password` en dur.**
- Idem pour `OPENAI_API_KEY` / `NOTION_API_KEY` : passer par `st.secrets`, plus de `.env` commité.
- Vérifier que `.streamlit/secrets.toml` et `.env` sont bien dans `.gitignore`.
- **Ne PAS écrire de vrais mots de passe dans le code ni dans ce repo.** Laisser des placeholders.

**Critère d'acceptation F6 :** `grep -i password email_automation_app.py` ne retourne plus aucun mot de passe en clair ; l'app charge les users depuis `st.secrets` ; l'app affiche un message clair si `secrets.toml` est absent.

---

## 4. Exigences non fonctionnelles

- Ne casser aucune fonctionnalité existante (mode test, reprise via `progress.json`, CC, pièces jointes, images inline, placeholders).
- Pas de nouvelle dépendance lourde. `random`, `json`, `os`, `datetime` sont en standard. (Si besoin d'un check MX un jour : `dnspython`, mais **hors périmètre** ici.)
- Le code doit rester lisible : factoriser l'envoi unitaire plutôt que dupliquer (voir §6).

---

## 5. Hors périmètre (à NE PAS faire dans ce ticket)

- Configuration DNS SPF/DKIM/DMARC (fait par Taddeo côté DNS + console Workspace).
- Création/usage d'un sous-domaine dédié.
- Import automatique des bounces via IMAP (phase ultérieure).
- Endpoint HTTP de désinscription one-click (phase ultérieure ; pour l'instant `mailto:` seulement).
- Double opt-in, base de données, dashboard de logs.

---

## 6. Refactor recommandé (pour éviter la double maintenance)

Extraire une fonction :

```python
def send_one_email(server, sender_email, email_data, cc_emails, automation, session_assets) -> bool:
    """Construit le MIME (mixed/alternative/related), pose les en-têtes List-Unsubscribe/Reply-To,
    attache images inline + pièces jointes, et appelle server.sendmail.
    Retourne True si envoyé, False sinon. Ne gère PAS le sleep ni le cap (gérés par l'appelant)."""
```

Les deux boucles appellent cette fonction. Les vérifications `is_suppressed()` et `get_today_count()` ainsi que `time.sleep(random.uniform(1,10))` et `incr_today_count()` restent dans l'appelant (la boucle), car elles pilotent le flux.

---

## 7. Checklist d'acceptation globale

- [ ] F1 : 3 en-têtes présents sur un email de test (vérifier le `.eml` source).
- [ ] F2 : lien de désinscription visible en HTML **et** en texte.
- [ ] F3 : un email suppressed n'est jamais envoyé (testé sur les 2 boucles) + UI d'ajout fonctionnelle.
- [ ] F4 : intervalle aléatoire 1–10 s, plus de `sleep(1)` fixe.
- [ ] F5 : plafond quotidien par expéditeur, persistant le jour même, configurable.
- [ ] F6 : zéro mot de passe en clair ; chargement via `st.secrets` ; `.gitignore` à jour.
- [ ] Mode test, reprise progress.json, CC, pièces jointes, images inline : toujours fonctionnels.
- [ ] L'app démarre sans erreur et envoie correctement un lot de test de 5 emails vers soi-même.
