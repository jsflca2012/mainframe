import os

# ===============
# Django Settings
# ===============

# much secure
DATABASE_PASSWORD = os.environ.get("DATABASE_PASSWORD")

DEBUG = False

ALLOWED_HOSTS = [
    "ip-172-31-18-22.us-west-2.compute.internal",
    "api.omoulearning.net",
    "loadbalancer-fd2287fde68b401a.elb.us-west-2.amazonaws.com",
    "3.19.70.245",
    "localhost",
]

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.postgresql_psycopg2",
        "NAME": "mainframe",
        "USER": "postgres",
        "PASSWORD": DATABASE_PASSWORD,
        "HOST": "mainframe.crjrqgmavbsy.us-west-2.rds.amazonaws.com",
        "PORT": "5432",
    },
}
SECRET_KEY = os.environ.get("SECRET_KEY")
SENDGRID_API_KEY = os.environ.get("SENDGRID_API_KEY")
