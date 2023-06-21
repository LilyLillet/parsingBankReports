class tinkoffOperation:
    def __init__(self, date, operationType, category, payment, comment):
        self.date = date
        self.operationType = operationType
        self.category = category
        self.payment = payment
        self.comment = comment

    def __str__(self):
        print(str(self.date) + ' ' + self.operationType + ' ' + self.category + ' ' + str(self.payment) + ' ' + self.comment)
