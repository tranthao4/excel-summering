# -*- coding: utf-8 -*-
"""
Konfigurationsfil för Excel Sammanslagning
Här kan du enkelt lägga till, ta bort eller ändra kolumner
"""

# Kolumndefinitioner
# Format: {
#   'id': Unikt ID för kolumnen (används internt)
#   'label': Text som visas i UI
#   'output_name': Kolumnnamn i summeringsfilen
#   'keywords': Lista med nyckelord för automatisk igenkänning (case-insensitive)
#   'required': True om kolumnen är obligatorisk, False om valfri
#   'help_text': Hjälptext som visas bredvid (valfritt)
#   'validate_content': Funktion för att validera innehåll (valfritt)
# }

COLUMNS = [
    {
        'id': 'unit',
        'label': 'Enhet:',
        'output_name': 'Enhet',
        'keywords': ['enhet', 'unit', 'avdelning', 'department', 'bolagsnamn'],
        'required': False,
        'help_text': None,
        'exclude_keywords': []
    },
    {
        'id': 'name',
        'label': 'Namn:',
        'output_name': 'Namn',
        'keywords': ['namn', 'name', 'förnamn', 'firstname', 'konsult'],
        'required': True,
        'help_text': None,
        'exclude_keywords': ['alternativ']  # Undvik kolumner med dessa ord
    },
    {
        'id': 'lastname',
        'label': 'Efternamn:',
        'output_name': None,  # Kombineras med Namn
        'keywords': ['efternamn', 'lastname', 'surname'],
        'required': False,
        'help_text': '(Kombineras med Namn)',
        'exclude_keywords': []
    },
    {
        'id': 'personnr',
        'label': 'Personnummer:',
        'output_name': 'Personnummer',
        'keywords': ['personnummer', 'personnr', 'ssn', 'social security'],
        'required': True,
        'help_text': None,
        'validate_content': 'personnummer',  # Speciell validering
        'exclude_keywords': []
    },
    {
        'id': 'address',
        'label': 'Adress:',
        'output_name': 'Adress',
        'keywords': ['arbetsställe', 'workplace', 'arbetsstalle'],
        'required': False,
        'help_text': None,
        'exclude_keywords': ['leverans', 'postort', 'post']
    },
    {
        'id': 'date_from',
        'label': 'Anställd från datum:',
        'output_name': 'Anställd from',
        'keywords': ['från datum', 'from date', 'startdatum', 'start date', 'anställd från', 'första arbetsdag'],
        'required': False,
        'help_text': None,
        'exclude_keywords': []
    },
    {
        'id': 'date_to',
        'label': 'Anställd till datum:',
        'output_name': 'Anställd tom',
        'keywords': ['till datum', 'to date', 'slutdatum', 'end date', 'anställd till'],
        'required': False,
        'help_text': '(Beräknar till idag om inget anges)',
        'exclude_keywords': []
    },
    {
        'id': 'days',
        'label': 'Anställningsdagar:',
        'output_name': 'Anställningsdagar',
        'keywords': ['arbetsdagar', 'anställningstid', 'antal dagar'],
        'required': False,
        'help_text': '(Beräknas automatiskt från anställningsdatum till idag om inget anges)',
        'exclude_keywords': []
    },
]

# Kolumner i summeringsfilen (i denna ordning)
SUMMARY_COLUMNS = ['Enhet', 'Namn', 'Adress', 'Personnummer', 'Anställd from', 'Anställd tom', 'Anställningsdagar']

# Valideringsregler
VALIDATION_RULES = {
    'personnummer': {
        'min_digits': 10,  # Minst 10 siffror
        'match_threshold': 0.8  # 80% av raderna måste matcha
    }
}

# UI-inställningar
UI_SETTINGS = {
    'window_title': 'Michelles magic tool',
    'window_size': '850x600',
    'clear_button_text': '✕',
    'clear_button_width': 2,
    # PostNord TPL färger
    'primary_color': '#00A0DC',  # PostNord blå
    'secondary_color': '#0088CC',  # Mörkare blå
    'background_color': '#FFFFFF',  # Vit bakgrund
    'text_color': '#333333',  # Mörkgrå text
    'logo_text': 'postnord',  # Logo text
    'logo_subtext': 'TPL'  # Logo subtext
}

# Filnamn
FILES = {
    'summary_file': 'summeringsfil.xlsx'
}
