# Generated by Django 2.2.4 on 2019-08-17 20:55

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('account', '0009_auto_20190728_1915'),
    ]

    operations = [
        migrations.AddField(
            model_name='admin',
            name='birth_date',
            field=models.DateField(default=datetime.date(2019, 8, 17)),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='instructor',
            name='birth_date',
            field=models.DateField(default=datetime.date(2019, 8, 17)),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='parent',
            name='birth_date',
            field=models.DateField(default=datetime.date(2019, 8, 17)),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='student',
            name='birth_date',
            field=models.DateField(default=datetime.date(2019, 8, 17)),
            preserve_default=False,
        ),
    ]
