import dataframe
import mail

"""
nom_client = "X"
numero_arbre = 5
nom_fr_arbre = "Chêne"
x = 10.22
y = 42.69
date_releve = "06/01/2020"
type_ope = "intervention"
date_limite = "06/04/2021"

text = "Cher Monsieur, Madame %s, \
    \n\nNous nous permettons de vous relancer concernant l’arbre n°%d, %s, situé aux coordonnées [%f, %f].\
    \nNous sommes intervenus sur cet arbre le %s, pour %s.\
    \nNous souhaitons effectuer une nouvelle visite pour %s sur cet arbre dans les prochains mois.\
    \n\nNous vous prions de nous recontacter pour confirmer cette nouvelle action à effectuer au plus tard le %s.\
    \n\nBien cordialement,\n\nONF" % (
        nom_client, 
        numero_arbre, 
        nom_fr_arbre, 
        x, 
        y, 
        date_releve, 
        type_ope, 
        type_ope, 
        date_limite)
"""

df = dataframe.dataframe("Excel_type.xlsx")
df.convert()
df_return = df.get_info_user_relance()

email = mail.mail(df_return)

for i in email.email_messages:
    print(i)