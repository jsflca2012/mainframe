# Generated by Django 2.2.8 on 2019-12-28 05:20

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('pricing', '0008_auto_20191228_0520'),
    ]

    operations = [
        migrations.AlterField(
            model_name='discount',
            name='amount',
            field=models.DecimalField(decimal_places=2, default=0.0, max_digits=5),
        ),
    ]
