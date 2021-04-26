# clients_info_preprocessing

 一个简单的客户信息预处理工具
 @ArmandXUuu 11 avr 2021
 Rosière-prés-Troyes

## Introduction des trois modes

1. `-n`
Mode normal, le programme cherchera le file `input.xlsx` dans le répertoire.

```python
python3 main.py -n
```

2. `-b`
Mode batch, le programme va traiter les fichiers suivant la commande.

```python
python3 main.py -b file1.xlsx file2.xlsx ...
```

3. `-ib`
Mode interactive et batch, recommandé, vous allez devoir voir s'il y a un espace entre le nom du client.

```python
python3 main.py -ib file1.xlsx file2.xlsx ...
```

## Attention

- Comme il n'y a aucun moyen de connaître la longueur et le nombre de mots du nom du client, deux `buffer` sont définis pour éviter de perdre le nom complet du client.
