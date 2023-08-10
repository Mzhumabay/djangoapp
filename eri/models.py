from django.db import models

# Create your models here.
from django.db import models

class RawQueryDictionary(models.Model):
    raw_queries = models.JSONField()

    def __str__(self):
        return f"Raw Query Dictionary (ID: {self.id})"
