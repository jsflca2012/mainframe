# Generated by Django 2.2.12 on 2020-07-25 05:50

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('account', '0040_auto_20200722_0128'),
        ('comms', '0005_parentnotificationsettings'),
    ]

    operations = [
        migrations.CreateModel(
            name='InstructorNotificationSettings',
            fields=[
                ('instructor', models.OneToOneField(on_delete=django.db.models.deletion.PROTECT, primary_key=True, serialize=False, to='account.Instructor')),
                ('session_reminder_email', models.BooleanField(default=True)),
                ('session_reminder_sms', models.BooleanField(default=False)),
                ('schedule_updates_sms', models.BooleanField(default=False)),
                ('course_requests_sms', models.BooleanField(default=False)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
            ],
        ),
        migrations.RemoveField(
            model_name='parentnotificationsettings',
            name='course_requests_sms',
        ),
    ]