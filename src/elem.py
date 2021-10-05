class Organization:

    def __init__(self,
                 name,
                 name_abbreviation,
                 address,
                 postcode,
                 city):
        self.name = name
        self.name_abbreviation = name_abbreviation
        self.address = address
        self.postcode = postcode
        self.city = city

class Person:

    def __init__(self,
                 name = None,
                 gender = None,
                 address = None,
                 postcode = None,
                 city = None,
                 telephone_number = None,
                 email_address = None):
        self.name = name
        if self.name is not None:
            self.surname = name.split()[-1]
        else:
            self.surname = None
        self.gender = gender
        self.address = address
        self.postcode = postcode
        self.city = city
        self.telephone_number = telephone_number
        self.email_address = email_address
