# Generated by Django 3.0.3 on 2020-07-25 04:55

from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('course', '0027_auto_20200720_0339'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('comms', '0005_parentnotificationsettings'),
    ]

    operations = [
        migrations.CreateModel(
            name='Annoucement',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('timestamp', models.DateTimeField(auto_now_add=True)),
                ('subject', models.TextField(blank=True)),
                ('body', models.TextField()),
                ('course', models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, to='course.Course')),
                ('user', models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, to=settings.AUTH_USER_MODEL)),
            ],
        ),
    ]
