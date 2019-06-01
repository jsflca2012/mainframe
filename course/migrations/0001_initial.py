# Generated by Django 2.2.1 on 2019-05-25 20:41

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        ('account', '0003_instructor'),
    ]

    operations = [
        migrations.CreateModel(
            name='Course',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('type', models.CharField(choices=[('C', 'Class'), ('T', 'Tutoring')], default='C', max_length=1)),
                ('subject', models.CharField(max_length=100)),
                ('description', models.CharField(max_length=1000)),
                ('room', models.CharField(max_length=50)),
                ('days', models.CharField(max_length=10)),
                ('schedule', models.TimeField()),
                ('start_date', models.DateField()),
                ('end_date', models.DateField()),
                ('instructor', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='account.Instructor')),
            ],
        ),
    ]
