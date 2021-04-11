# clients_info_preprocessing

 一个简单的客户信息预处理工具
 @ArmandXUuu 11 avr 2021
 Rosière-prés-Troyes

## Pour commencer

1. Renommez le fichier `xxx.xlsx` en `input.csv`

2. Exécutez la commande

```bash
python3 main.py
```

3. Done ! Le résultat se trouve dans le même répertoire nommé `output.csv`

## Attention

- Comme il n'y a aucun moyen de connaître la longueur et le nombre de mots du nom du client, deux `buffer` sont définis pour éviter de perdre le nom complet du client.
