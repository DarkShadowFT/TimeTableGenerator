class Period:
	def __init__(self, startTime, endTime, name, section, venue, day):
		self.startTime = startTime
		self.endTime = endTime
		self.name = name
		self.section = section
		self.venue = venue
		self.day = day

	def __str__(self):
		return f"{self.name}({self.section})\t\t{self.startTime}-{self.endTime}\t{self.venue}\t\t{self.day}"