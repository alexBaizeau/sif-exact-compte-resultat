# Exact online api demo
## Instalation steps

```bash

git clone https://github.com/alexBaizeau/exact-online-api-demo.git
cd exact-online-api-demo
virtualenv env
. ./env/bin/activate
pip install -r requirements.txt
cp config.ini.sample config.ini
```


## First time user
Make sure that you have an app here https://apps.exactonline.com/be/fr-BE/Manage if not create a test one
That's where the client id , secret and url are

```bash
python compte_de_resultat.py setup --base-url=https://www.mycompany.com --client-id={XXXXXX-xxxx-xxxx-xxxx-XXXXXXXX} --client-secret=XXXXX
```

## Excel

```bash
python compte_de_resultat.py excel --annee_fiscale=2018
```

## Create an executable avec PyInstaller
```
./env/bin/pyinstaller -p ./env/bin/python compte_de_resultat.py --add-data 'config.ini:.' --add-data 'rapport_config.json:.'
```
