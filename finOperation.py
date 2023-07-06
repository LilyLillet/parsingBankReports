class finOperation:
    def __init__(self, date, operation_type, counterparty, payment, comment):
        self.date = date
        self.operationType = operation_type
        self.counterparty = counterparty
        self.payment = payment
        self.comment = comment

    def __str__(self):
        return str(self.date) + ' ' + self.operationType + ' ' + self.counterparty + ' ' + str(
            self.payment) + ' ' + self.comment
