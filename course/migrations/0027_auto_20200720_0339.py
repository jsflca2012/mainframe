# Generated by Django 2.2.12 on 2020-07-20 03:39

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('course', '0026_auto_20200524_2101'),
    ]

    operations = [
        migrations.AlterField(
            model_name='coursecategory',
            name='description',
            field=models.CharField(blank=True, max_length=1000),
        ),
    ]