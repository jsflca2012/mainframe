# Generated by Django 2.2.3 on 2019-12-29 21:55

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('course', '0018_auto_20191215_0058'),
    ]

    operations = [
        migrations.RenameField(
            model_name='course',
            old_name='tuition',
            new_name='hourly_tuition',
        ),
        migrations.AddField(
            model_name='course',
            name='academic_level',
            field=models.CharField(blank=True, choices=[('elementary', 'Elementary School Level'), ('middle', 'Middle School Level'), ('high', 'High School Level'), ('college', 'College Level')], max_length=15, null=True),
        ),
        migrations.AddField(
            model_name='course',
            name='total_tuition',
            field=models.DecimalField(blank=True, decimal_places=2, max_digits=6, null=True),
        ),
        migrations.AddField(
            model_name='enrollment',
            name='num_sessions',
            field=models.IntegerField(default=0),
        ),
        migrations.AlterField(
            model_name='course',
            name='type',
            field=models.CharField(choices=[('C', 'Class'), ('S', 'Small Group'), ('T', 'Tutoring')], default='C', max_length=1),
        ),
    ]