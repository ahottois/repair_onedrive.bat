@echo off
setlocal enabledelayedexpansion
title Utilitaire de Réparation OneDrive

:: Exécuter en tant qu'administrateur
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Demande de privilèges administrateur...
    goto UACPrompt
) else (
    goto gotAdmin
)

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs"
    pushd "%CD%"
    CD /D "%~dp0"

:menu
cls
echo.
echo ============================================
echo      UTILITAIRE DE RÉPARATION ONEDRIVE
echo ============================================
echo.
echo  1. Vérifier l'état de OneDrive
echo  2. Déconnecter OneDrive
echo  3. Réinitialiser OneDrive
echo  4. Réinstaller OneDrive
echo  5. Corriger les erreurs courantes
echo  6. Réparer les permissions
echo  7. Effacer le cache de OneDrive
echo  8. Quitter
echo.
echo ============================================
echo.

set /p choix="Entrez votre choix (1-8): "

if "%choix%"=="1" goto check_status
if "%choix%"=="2" goto disconnect
if "%choix%"=="3" goto reset
if "%choix%"=="4" goto reinstall
if "%choix%"=="5" goto fix_common
if "%choix%"=="6" goto fix_permissions
if "%choix%"=="7" goto clear_cache
if "%choix%"=="8" goto exit_script

echo Choix invalide. Veuillez réessayer.
timeout /t 2 >nul
goto menu

:check_status
cls
echo.
echo ============================================
echo       VÉRIFICATION DE L'ÉTAT ONEDRIVE
echo ============================================
echo.

:: Vérifier si OneDrive est en cours d'exécution
tasklist /fi "imagename eq OneDrive.exe" | find /i "OneDrive.exe" >nul
if %errorlevel% equ 0 (
    echo [+] OneDrive est en cours d'exécution.
) else (
    echo [-] OneDrive n'est pas en cours d'exécution.
)

:: Vérifier l'emplacement de OneDrive
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    echo [+] OneDrive est installé à l'emplacement par défaut.
) else (
    echo [-] OneDrive n'est pas installé à l'emplacement par défaut.
)

:: Vérifier la connexion Internet
ping -n 1 onedrive.live.com >nul
if %errorlevel% equ 0 (
    echo [+] La connexion à OneDrive.live.com est établie.
) else (
    echo [-] Impossible de se connecter à OneDrive.live.com.
)

:: Vérifier si OneDrive démarre avec Windows
reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive >nul 2>&1
if %errorlevel% equ 0 (
    echo [+] OneDrive est configuré pour démarrer avec Windows.
) else (
    echo [-] OneDrive n'est pas configuré pour démarrer avec Windows.
)

echo.
pause
goto menu

:disconnect
cls
echo.
echo ============================================
echo         DÉCONNEXION DE ONEDRIVE
echo ============================================
echo.
echo Fermeture de OneDrive...
taskkill /f /im OneDrive.exe >nul 2>&1

echo Déconnexion de OneDrive...
echo Cette opération va déconnecter vos comptes OneDrive sans supprimer vos fichiers.
echo.
set /p confirm="Êtes-vous sûr de vouloir continuer? (O/N): "
if /i "%confirm%"=="O" (
    echo.
    echo Déconnexion de OneDrive en cours...
    
    :: Déconnecter OneDrive
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /shutdown
        "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /unlink
        echo Comptes OneDrive déconnectés avec succès.
    ) else (
        echo OneDrive n'est pas installé à l'emplacement par défaut.
    )
) else (
    echo Opération annulée.
)

echo.
pause
goto menu

:reset
cls
echo.
echo ============================================
echo         RÉINITIALISATION DE ONEDRIVE
echo ============================================
echo.
echo Cette opération va réinitialiser OneDrive et supprimer tous les paramètres.
echo Vos fichiers ne seront pas supprimés, mais vous devrez reconfigurer OneDrive.
echo.
set /p confirm="Êtes-vous sûr de vouloir continuer? (O/N): "
if /i "%confirm%"=="O" (
    echo.
    echo Réinitialisation de OneDrive en cours...
    
    :: Arrêter OneDrive
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    :: Réinitialiser OneDrive
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /reset
        echo OneDrive a été réinitialisé avec succès.
        echo Redémarrage de OneDrive...
        start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    ) else (
        echo OneDrive n'est pas installé à l'emplacement par défaut.
    )
) else (
    echo Opération annulée.
)

echo.
pause
goto menu

:reinstall
cls
echo.
echo ============================================
echo         RÉINSTALLATION DE ONEDRIVE
echo ============================================
echo.
echo Cette opération va complètement désinstaller puis réinstaller OneDrive.
echo Vos fichiers ne seront pas supprimés, mais vous devrez reconfigurer OneDrive.
echo.
set /p confirm="Êtes-vous sûr de vouloir continuer? (O/N): "
if /i "%confirm%"=="O" (
    echo.
    echo Désinstallation de OneDrive en cours...
    
    :: Arrêter OneDrive
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    :: Désinstaller OneDrive
    if exist "%SYSTEMROOT%\SysWOW64\OneDriveSetup.exe" (
        %SYSTEMROOT%\SysWOW64\OneDriveSetup.exe /uninstall
    ) else (
        if exist "%SYSTEMROOT%\System32\OneDriveSetup.exe" (
            %SYSTEMROOT%\System32\OneDriveSetup.exe /uninstall
        ) else (
            echo OneDrive n'est pas installé à l'emplacement par défaut.
            goto:reinstall_end
        )
    )
    
    echo OneDrive a été désinstallé avec succès.
    timeout /t 10 >nul
    
    echo Réinstallation de OneDrive en cours...
    :: Réinstaller OneDrive
    if exist "%SYSTEMROOT%\SysWOW64\OneDriveSetup.exe" (
        %SYSTEMROOT%\SysWOW64\OneDriveSetup.exe
    ) else (
        if exist "%SYSTEMROOT%\System32\OneDriveSetup.exe" (
            %SYSTEMROOT%\System32\OneDriveSetup.exe
        ) else (
            echo Impossible de trouver le programme d'installation de OneDrive.
            goto:reinstall_end
        )
    )
    
    echo OneDrive a été réinstallé avec succès.
) else (
    echo Opération annulée.
)

:reinstall_end
echo.
pause
goto menu

:fix_common
cls
echo.
echo ============================================
echo     CORRECTION DES ERREURS COURANTES
echo ============================================
echo.
echo Choisissez l'erreur à corriger:
echo.
echo  1. OneDrive ne démarre pas
echo  2. Erreurs de synchronisation
echo  3. Fichiers en conflit
echo  4. Problèmes de connexion
echo  5. OneDrive utilise trop de ressources
echo  6. Retour au menu principal
echo.
echo ============================================
echo.

set /p error_choice="Entrez votre choix (1-6): "

if "%error_choice%"=="1" goto fix_startup
if "%error_choice%"=="2" goto fix_sync
if "%error_choice%"=="3" goto fix_conflicts
if "%error_choice%"=="4" goto fix_connection
if "%error_choice%"=="5" goto fix_resources
if "%error_choice%"=="6" goto menu

echo Choix invalide. Veuillez réessayer.
timeout /t 2 >nul
goto fix_common

:fix_startup
cls
echo.
echo ============================================
echo     CORRECTION - ONEDRIVE NE DÉMARRE PAS
echo ============================================
echo.
echo Exécution des corrections...

:: Arrêter OneDrive s'il est en cours d'exécution
taskkill /f /im OneDrive.exe >nul 2>&1

:: Réenregistrer les DLL de OneDrive
echo 1/5 - Réenregistrement des DLL...
regsvr32 /s "%SYSTEMROOT%\System32\shell32.dll"

:: Vérifier le démarrage automatique
echo 2/5 - Vérification du démarrage automatique...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v OneDrive /t REG_SZ /d "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /f >nul 2>&1

:: Réinitialiser OneDrive
echo 3/5 - Réinitialisation de OneDrive...
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" /reset >nul 2>&1
)

:: Nettoyer le registre
echo 4/5 - Nettoyage du registre...
reg delete "HKCU\Software\Microsoft\OneDrive" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Policies\Microsoft\OneDrive" /f >nul 2>&1

:: Redémarrer OneDrive
echo 5/5 - Redémarrage de OneDrive...
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive a été redémarré.
) else (
    echo OneDrive n'est pas installé à l'emplacement par défaut.
)

echo.
echo Corrections terminées. Vérifiez si OneDrive démarre correctement maintenant.
echo.
pause
goto fix_common

:fix_sync
cls
echo.
echo ============================================
echo    CORRECTION - ERREURS DE SYNCHRONISATION
echo ============================================
echo.
echo Exécution des corrections...

:: Arrêter OneDrive
taskkill /f /im OneDrive.exe >nul 2>&1

:: Effacer le cache de synchronisation
echo 1/4 - Effacement du cache de synchronisation...
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Personal" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Personal\*.* >nul 2>&1
)
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Business1" (
    del /f /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\settings\Business1\*.* >nul 2>&1
)

:: Réinitialiser l'ODSP
echo 2/4 - Réinitialisation des caches ODSP...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" (
    rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" >nul 2>&1
)

:: Vérifier les listes d'exclusion
echo 3/4 - Vérification des listes d'exclusion...
reg delete "HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\Personal\Excluded" /f >nul 2>&1
reg delete "HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Excluded" /f >nul 2>&1

:: Redémarrer OneDrive
echo 4/4 - Redémarrage de OneDrive...
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive a été redémarré.
) else (
    echo OneDrive n'est pas installé à l'emplacement par défaut.
)

echo.
echo Corrections terminées. Vérifiez si les problèmes de synchronisation sont résolus.
echo Notez que la première synchronisation après cette opération peut prendre du temps.
echo.
pause
goto fix_common

:fix_conflicts
cls
echo.
echo ============================================
echo      CORRECTION - FICHIERS EN CONFLIT
echo ============================================
echo.
echo Cette opération va vous aider à résoudre les problèmes de fichiers en conflit.
echo.

:: Vérifier les dossiers de conflit
echo Vérification des dossiers de conflit...
if exist "%USERPROFILE%\OneDrive\*conflict*" (
    echo Des fichiers en conflit ont été trouvés dans votre dossier OneDrive.
) else (
    echo Aucun fichier en conflit n'a été détecté.
)

echo.
echo Conseils pour résoudre les conflits:
echo 1. Ouvrez votre dossier OneDrive et recherchez les fichiers avec "conflit" dans leur nom.
echo 2. Comparez les versions et conservez celle qui contient les modifications les plus récentes.
echo 3. Supprimez les fichiers en conflit une fois que vous avez sauvegardé les données importantes.
echo.
echo Pour éviter les conflits à l'avenir:
echo - Évitez de modifier le même fichier simultanément sur plusieurs appareils.
echo - Assurez-vous que OneDrive a terminé la synchronisation avant d'éteindre votre ordinateur.
echo - Vérifiez régulièrement l'état de synchronisation de OneDrive.
echo.

pause
goto fix_common

:fix_connection
cls
echo.
echo ============================================
echo      CORRECTION - PROBLÈMES DE CONNEXION
echo ============================================
echo.
echo Exécution des corrections...

:: Vérifier la connexion réseau
echo 1/5 - Vérification de la connexion réseau...
ping -n 1 onedrive.live.com >nul
if %errorlevel% equ 0 (
    echo    La connexion à OneDrive.live.com est établie.
) else (
    echo    Impossible de se connecter à OneDrive.live.com.
    echo    Vérifiez votre connexion Internet et réessayez.
    echo.
    pause
    goto fix_common
)

:: Réinitialiser Winsock et IP
echo 2/5 - Réinitialisation de Winsock...
netsh winsock reset >nul 2>&1

echo 3/5 - Réinitialisation de la pile TCP/IP...
netsh int ip reset >nul 2>&1

:: Vider le cache DNS
echo 4/5 - Vidage du cache DNS...
ipconfig /flushdns >nul 2>&1

:: Redémarrer OneDrive
echo 5/5 - Redémarrage de OneDrive...
taskkill /f /im OneDrive.exe >nul 2>&1
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive a été redémarré.
) else (
    echo OneDrive n'est pas installé à l'emplacement par défaut.
)

echo.
echo Corrections terminées. Vérifiez si les problèmes de connexion sont résolus.
echo Pour que les changements prennent effet complètement, il est recommandé de redémarrer votre ordinateur.
echo.
set /p reboot="Voulez-vous redémarrer votre ordinateur maintenant? (O/N): "
if /i "%reboot%"=="O" (
    shutdown /r /t 60 /c "Redémarrage planifié pour appliquer les modifications réseau."
    echo L'ordinateur va redémarrer dans 60 secondes. Fermez toutes vos applications.
    echo Pour annuler, tapez 'shutdown /a' dans une nouvelle fenêtre de commande.
)

echo.
pause
goto fix_common

:fix_resources
cls
echo.
echo ============================================
echo  CORRECTION - UTILISATION EXCESSIVE DE CPU
echo ============================================
echo.
echo Exécution des corrections...

:: Arrêter OneDrive
taskkill /f /im OneDrive.exe >nul 2>&1

:: Limiter l'utilisation de la bande passante
echo 1/3 - Ajustement des paramètres de bande passante...
reg add "HKCU\SOFTWARE\Microsoft\OneDrive\Accounts\Personal" /v "DiskSpaceCheckThresholdMB" /t REG_DWORD /d 1024 /f >nul 2>&1
reg add "HKCU\SOFTWARE\Microsoft\OneDrive" /v "MaxBandwidthKB" /t REG_DWORD /d 512 /f >nul 2>&1

:: Ajuster les paramètres de performance
echo 2/3 - Optimisation des paramètres de performance...
reg add "HKCU\SOFTWARE\Microsoft\OneDrive" /v "ProcessResourcePriority" /t REG_DWORD /d 1 /f >nul 2>&1

:: Redémarrer OneDrive
echo 3/3 - Redémarrage de OneDrive...
timeout /t 5 >nul
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
    start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
    echo OneDrive a été redémarré avec des paramètres optimisés.
) else (
    echo OneDrive n'est pas installé à l'emplacement par défaut.
)

echo.
echo Corrections terminées. Vérifiez si l'utilisation des ressources est améliorée.
echo.
echo Conseils supplémentaires:
echo - Réduisez le nombre de fichiers synchronisés en choisissant des dossiers spécifiques.
echo - Excluez les fichiers volumineux ou temporaires de la synchronisation.
echo - Mettez à jour OneDrive vers la dernière version disponible.
echo.

pause
goto fix_common

:fix_permissions
cls
echo.
echo ============================================
echo        RÉPARATION DES PERMISSIONS
echo ============================================
echo.
echo Cette opération va corriger les problèmes de permissions de OneDrive.
echo.
set /p confirm="Êtes-vous sûr de vouloir continuer? (O/N): "
if /i "%confirm%"=="O" (
    echo.
    echo Réparation des permissions en cours...
    
    :: Arrêter OneDrive
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    :: Réparer les permissions OneDrive
    echo 1/3 - Correction des permissions du dossier OneDrive...
    icacls "%USERPROFILE%\OneDrive" /reset /T /C /Q >nul 2>&1
    
    :: Réparer les permissions des clés de registre
    echo 2/3 - Correction des permissions du registre...
    reg add "HKCU\Software\Microsoft\OneDrive" /f >nul 2>&1
    
    :: Redémarrer OneDrive
    echo 3/3 - Redémarrage de OneDrive...
    timeout /t 5 >nul
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
        echo OneDrive a été redémarré.
    ) else (
        echo OneDrive n'est pas installé à l'emplacement par défaut.
    )
    
    echo Réparation des permissions terminée.
) else (
    echo Opération annulée.
)

echo.
pause
goto menu

:clear_cache
cls
echo.
echo ============================================
echo        EFFACEMENT DU CACHE ONEDRIVE
echo ============================================
echo.
echo Cette opération va effacer tous les fichiers de cache de OneDrive.
echo Cela peut résoudre de nombreux problèmes de synchronisation, mais la
echo première synchronisation après cette opération peut prendre du temps.
echo.
set /p confirm="Êtes-vous sûr de vouloir continuer? (O/N): "
if /i "%confirm%"=="O" (
    echo.
    echo Effacement du cache en cours...
    
    :: Arrêter OneDrive
    taskkill /f /im OneDrive.exe >nul 2>&1
    
    :: Effacer le cache
    echo 1/3 - Suppression des fichiers de cache...
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\cache" (
        rmdir /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\cache" >nul 2>&1
    )
    
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\settings" (
        rmdir /s /q "%LOCALAPPDATA%\Microsoft\OneDrive\settings" >nul 2>&1
    )
    
    if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" (
        rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" >nul 2>&1
    )
    
    :: Recréer les dossiers
    echo 2/3 - Recréation des dossiers de cache...
    mkdir "%LOCALAPPDATA%\Microsoft\OneDrive\cache" >nul 2>&1
    mkdir "%LOCALAPPDATA%\Microsoft\OneDrive\settings" >nul 2>&1
    
    :: Redémarrer OneDrive
    echo 3/3 - Redémarrage de OneDrive...
    timeout /t 5 >nul
    if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
        start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
        echo OneDrive a été redémarré.
    ) else (
        echo OneDrive n'est pas installé à l'emplacement par défaut.
    )
    
    echo Effacement du cache terminé.
) else (
    echo Opération annulée.
)

echo.
pause
goto menu

:exit_script
cls
echo.
echo ============================================
echo      MERCI D'AVOIR UTILISÉ L'UTILITAIRE
echo ============================================
echo.
echo À bientôt!
echo.
timeout /t 3 >nul
exit
