# Generated by Django 2.2.4 on 2019-08-20 05:16

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('account', '0010_instructornote_note_parentnote_studentnote'),
    ]

    operations = [
        migrations.AlterField(
            model_name='note',
            name='body',
            field=models.TextField(blank=True),
        ),
        migrations.AlterField(
            model_name='studentnote',
            name='body',
            field=models.TextField(blank=True),
        ),
    ]