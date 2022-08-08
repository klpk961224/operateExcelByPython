class StoreInfoObj(object):
    def __init__(self, store_name, store_url, store_address, store_consumption, store_score, comment_count):
        self._store_name = store_name
        self._store_url = store_url
        self._store_address = store_address
        self._store_consumption = store_consumption
        self._store_score = store_score
        self._comment_count = comment_count

    @property
    def store_name(self):
        return self._store_name

    @store_name.setter
    def store_name(self, store_name):
        self._store_name = store_name

    @store_name.deleter
    def store_name(self):
        self._store_name = ''

    @property
    def store_url(self):
        return self._store_url

    @store_url.setter
    def store_url(self, store_url):
        self._store_url = store_url

    @store_url.deleter
    def store_url(self):
        self._store_url = ''

    @property
    def store_address(self):
        return self._store_address

    @store_address.setter
    def store_address(self, store_address):
        self._store_address = store_address

    @store_address.deleter
    def store_address(self):
        self._store_address = ''

    @property
    def store_consumption(self):
        return self._store_consumption

    @store_consumption.setter
    def store_consumption(self, store_consumption):
        self._store_consumption = store_consumption

    @store_consumption.deleter
    def store_consumption(self):
        self._store_consumption = ''

    @property
    def store_score(self):
        return self._store_score

    @store_score.setter
    def store_score(self, store_score):
        self._store_score = store_score

    @store_score.deleter
    def store_score(self):
        self._store_score = ''

    @property
    def comment_count(self):
        return self._comment_count

    @comment_count.setter
    def comment_count(self, comment_count):
        self._comment_count = comment_count

    @comment_count.deleter
    def comment_count(self):
        self._comment_count = ''
