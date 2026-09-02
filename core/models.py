"""
core/models.py — Customer เป็นของกลาง ใช้ร่วมกันทุกโมดูลลูกค้า (cpall, และลูกค้าเจ้าถัดๆ ไป)
"""
from django.db import models


class Customer(models.Model):
    code = models.CharField(max_length=20, unique=True, verbose_name="รหัสลูกค้า")
    name_th = models.TextField(verbose_name="ชื่อลูกค้า")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = "customer"
        managed = False
        verbose_name = "ลูกค้า"
        verbose_name_plural = "ลูกค้า"

    def __str__(self):
        return f"{self.name_th} ({self.code})"
