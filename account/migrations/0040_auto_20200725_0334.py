# Generated by Django 2.2.14 on 2020-07-25 03:34

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('account', '0039_auto_20200602_0359'),
    ]

    operations = [
        migrations.AlterField(
            model_name='school',
            name='district',
            field=models.CharField(blank=True, default='', max_length=100),
            preserve_default=False,
        ),
    ]
