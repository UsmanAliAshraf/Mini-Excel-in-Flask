class Cell:
    def __init__(self, value='', font_size=16, bold=False, italic=False):
        self.value = value
        self.font_size = font_size
        self.bold = bold
        self.italic = italic

    def to_dict(self):
        return {
            'value': self.value,
            'font_size': self.font_size,
            'bold': self.bold,
            'italic': self.italic
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            value=data.get('value', ''),
            font_size=data.get('font_size', 16),
            bold=data.get('bold', False),
            italic=data.get('italic', False)
        )
