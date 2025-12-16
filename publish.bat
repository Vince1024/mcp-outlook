@echo off
REM Script de publication GitHub pour MCP-OUTLOOK
REM Ce script vous guide à travers les étapes de publication

echo ========================================
echo   MCP-OUTLOOK - Script de Publication
echo ========================================
echo.

REM Vérifier si Git est installé
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Git n'est pas installé ou pas dans le PATH
    echo Téléchargez Git depuis: https://git-scm.com/download/win
    pause
    exit /b 1
)

echo [OK] Git est installé
echo.

REM Vérifier si on est dans un repo Git
if not exist ".git" (
    echo [INFO] Initialisation du repository Git...
    git init
    echo [OK] Repository Git initialisé
) else (
    echo [OK] Repository Git déjà initialisé
)
echo.

REM Afficher le statut
echo [INFO] Statut actuel du repository:
git status --short
echo.

REM Demander confirmation
set /p confirm="Voulez-vous ajouter tous les fichiers et créer le commit initial? (O/N): "
if /i not "%confirm%"=="O" (
    echo Publication annulée.
    pause
    exit /b 0
)

REM Ajouter tous les fichiers
echo.
echo [INFO] Ajout de tous les fichiers...
git add .
echo [OK] Fichiers ajoutés

REM Créer le commit
echo.
echo [INFO] Création du commit initial...
git commit -m "Initial commit: MCP-OUTLOOK v1.0.0 - Ready for public release"
if errorlevel 1 (
    echo [AVERTISSEMENT] Aucun changement à commiter ou erreur
) else (
    echo [OK] Commit créé
)

REM Demander le nom d'utilisateur GitHub
echo.
echo ========================================
echo   Configuration du Remote GitHub
echo ========================================
echo.
set /p github_user="Entrez votre nom d'utilisateur GitHub: "

if "%github_user%"=="" (
    echo [ERREUR] Nom d'utilisateur requis
    pause
    exit /b 1
)

REM Vérifier si le remote existe déjà
git remote get-url origin >nul 2>&1
if not errorlevel 1 (
    echo [INFO] Remote 'origin' existe déjà
    git remote -v
    echo.
    set /p change_remote="Voulez-vous le changer? (O/N): "
    if /i "%change_remote%"=="O" (
        git remote remove origin
        git remote add origin https://github.com/%github_user%/mcp-outlook.git
        echo [OK] Remote mis à jour
    )
) else (
    git remote add origin https://github.com/%github_user%/mcp-outlook.git
    echo [OK] Remote ajouté: https://github.com/%github_user%/mcp-outlook
)

echo.
echo ========================================
echo   Prêt pour la Publication
echo ========================================
echo.
echo IMPORTANT: Avant de pousser le code, assurez-vous d'avoir:
echo   1. Créé le repository sur GitHub: https://github.com/new
echo      - Nom: mcp-outlook
echo      - Visibilité: Public
echo      - N'initialisez RIEN (pas de README, .gitignore, ou licence)
echo.
echo   2. Configuré vos identifiants Git (si pas déjà fait):
echo      git config --global user.name "Votre Nom"
echo      git config --global user.email "votre@email.com"
echo.
set /p push_now="Voulez-vous pousser le code maintenant? (O/N): "

if /i "%push_now%"=="O" (
    echo.
    echo [INFO] Renommage de la branche en 'main'...
    git branch -M main
    
    echo [INFO] Push vers GitHub...
    git push -u origin main
    
    if errorlevel 1 (
        echo.
        echo [ERREUR] Le push a échoué
        echo.
        echo Causes possibles:
        echo   - Le repository n'existe pas sur GitHub
        echo   - Problème d'authentification
        echo   - Pas de connexion internet
        echo.
        echo Solutions:
        echo   1. Créez le repository sur GitHub
        echo   2. Configurez vos identifiants Git
        echo   3. Réessayez avec: git push -u origin main
    ) else (
        echo.
        echo ========================================
        echo   Publication Réussie! 🎉
        echo ========================================
        echo.
        echo Votre code est maintenant sur GitHub:
        echo https://github.com/%github_user%/mcp-outlook
        echo.
        echo Prochaines étapes:
        echo   1. Créer une release v1.0.0 sur GitHub
        echo   2. Ajouter les topics (mcp, outlook, python, etc.)
        echo   3. Partager le projet!
        echo.
        echo Consultez PUBLISHING_GUIDE.md pour plus de détails.
    )
) else (
    echo.
    echo Publication annulée.
    echo.
    echo Pour pousser manuellement plus tard:
    echo   git branch -M main
    echo   git push -u origin main
)

echo.
pause

