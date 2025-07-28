class Cell:
    def __init__(self, value='', font_size=16, bold=False, italic=False, bg_color=None, text_color=None):
        self.value = value
        self.font_size = font_size
        self.bold = bold
        self.italic = italic
        self.bg_color = bg_color  # background color (hex or CSS)
        self.text_color = text_color  # text color (hex or CSS)

    def to_dict(self):
        return {
            'value': self.value,
            'font_size': self.font_size,
            'bold': self.bold,
            'italic': self.italic,
            'bg_color': self.bg_color,
            'text_color': self.text_color
        }

    @classmethod
    def from_dict(cls, data):
        return cls(
            value=data.get('value', ''),
            font_size=data.get('font_size', 16),
            bold=data.get('bold', False),
            italic=data.get('italic', False),
            bg_color=data.get('bg_color'),
            text_color=data.get('text_color')
        )
