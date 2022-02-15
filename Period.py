class Period:
	def __init__(self, startTime, endTime, name, section, venue, day):
		self._startTime = startTime
		self._endTime = endTime
		self._name = name
		self._section = section
		self._venue = venue
		self._day = day

	@property
	def duration(self):
		return f"{self._startTime}-{self._endTime}"

	@property
	def name(self):
		return self._name

	@property
	def section(self):
		return self._section

	@property
	def venue(self):
		return self._venue

	@property
	def day(self):
		return self._day

	def tt_format(self):
		print(f"{self.name}({self.section})\t\t{self.venue}\t\t{self._startTime}-{self._endTime}")

	def __str__(self):
		return f"{self.name}\t\t{self.section}\t\t{self.venue}\t\t{self._startTime}-{self._endTime}\t\t{self._day}"