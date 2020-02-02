# Generated by Django 2.2.7 on 2020-01-25 22:49

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('account', '0028_auto_20200103_2235'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='instructoravailability',
            name='id',
        ),
        migrations.AlterField(
            model_name='instructoravailability',
            name='instructor',
            field=models.OneToOneField(on_delete=django.db.models.deletion.PROTECT, primary_key=True, serialize=False, to='account.Instructor'),
        ),
    ]
