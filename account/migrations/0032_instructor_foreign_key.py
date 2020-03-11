# Generated by Django 2.2.10 on 2020-02-23 03:23

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('account', '0031_refactor_instructor_availability'),
    ]

    operations = [
        migrations.AlterField(
            model_name='instructoravailability',
            name='instructor',
            field=models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, to='account.Instructor'),
        ),
        migrations.AddField(
            model_name='instructoravailability',
            name='id',
            field=models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID'),
            preserve_default=False,
        ),
    ]