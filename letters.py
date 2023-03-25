from fpdf import FPDF

def classe(id):
    if id==None or id=='Sheet':
        return 'RIEN'
    else:
        id = id.split('.')
        classe = id[2][0]
        match classe:
            case 'T':
                niveau = 'Terminale '
            case 'P':
                niveau = 'Première '
            case 'S':
                niveau = 'Seconde '
        division = id[2][1:6]
        match division:
            case 'STMG1':
                division = 'STMG 1'
            case 'STMG2':
                division = 'STMG 2'
        return niveau + division

def prenom(id):
    if id==None or id=='Sheet':
        return 'RIEN'
    else:
        id = id.split('.')
        prenom = id[0]
        return prenom.capitalize()

def nom(id):
    if id==None or id=='Sheet':
        return 'RIEN'
    else:
        id = id.split('.')
        nom = id[1]
        return nom.upper()

def generate_letters(matches):
    page=0
    pdf = FPDF(orientation="portrait", format="A4")
    pdf.add_page()
    for match in matches:
        if page % 4 == 0:
            pdf.add_page()
        pdf.set_font('Times', '', 11)
        pdf.cell(40, 10, f'Très chère {prenom(match)} {nom(match)}, élève de {classe(match)},', 0, 1)
        pdf.cell(40, 10,
                 'Merci énormément pour ta participation! Vous êtiez plus de 150 à jouer le jeu, et à répondre aux question de Cupidéon.',
                 0, 1)
        pdf.cell(40, 10,
                 "En cette journée du 14 février 2023, nous avons l'honneur de présenter ta moitié perdue, qui est:", 0,
                 1)
        pdf.set_font('Times', 'B', 11)
        pdf.cell(40, 10, f"{prenom(matches[match])} {nom(matches[match])}, {classe(matches[match])}", 0, 1)
        pdf.set_font('Times', '', 11)
        pdf.cell(40, 10, "Nous te souhaitons une belle Saint-Valentin.", 0, 1)
        pdf.cell(40, 10, "------", 0, 1)
        page +=1
        if page % 4 == 0:
            pdf.add_page()
        pdf.set_font('Times', '', 11)
        pdf.cell(40, 10, f'Très cher {prenom(matches[match])} {nom(matches[match])}, élève de {classe(matches[match])},', 0, 1)
        pdf.cell(40, 10,
                 'Merci énormément pour ta participation! Vous êtiez plus de 150 à jouer le jeu, et à répondre aux question de Cupidéon.',
                 0, 1)
        pdf.cell(40, 10,
                 "En cette journée du 14 février 2023, nous avons l'honneur de présenter ta moitié perdue, qui est:", 0,
                 1)
        pdf.set_font('Times', 'B', 11)
        pdf.cell(40, 10, f"{prenom(match)} {nom(match)}, {classe(match)}", 0, 1)
        pdf.set_font('Times', '', 11)
        pdf.cell(40, 10, "Nous te souhaitons une belle Saint-Valentin.", 0, 1)
        pdf.cell(40, 10, "------", 0, 1)
        page += 1

    pdf.output("GFGd.pdf")

