"""
Django settings for mainframe project.

Generated by "django-admin startproject" using Django 2.2.1.

For more information on this file, see
https://docs.djangoproject.com/en/2.2/topics/settings/

For the full list of settings and their values, see
https://docs.djangoproject.com/en/2.2/ref/settings/
"""
import importlib
import os
import sys


PROJECT_ROOT = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(PROJECT_ROOT, "apps"))

# Build paths inside the project like this: os.path.join(BASE_DIR, ...)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

env_module_name = os.environ.get("DJANGO_ENV_MODULE")
if env_module_name is None:  # pragma: no cover
    print("DJANGO_ENV_MODULE is not set!", file=sys.stderr)
    sys.exit(1)

cwd = os.getcwd()
env = None

if cwd in sys.path:
    env = importlib.import_module(env_module_name)
else:  # pragma: no cover
    sys.path.insert(0, cwd)
    try:
        env = importlib.import_module(env_module_name)
    finally:
        try:
            sys.path.remove(cwd)
        except ValueError:
            pass


DEBUG = env.DEBUG

if not DEBUG or hasattr(env, "ALLOWED_HOSTS"):
    ALLOWED_HOSTS = env.ALLOWED_HOSTS


# Application definition

API_APPS = (
    "account.apps.AccountConfig",
    "course.apps.CourseConfig",
    "scheduler.apps.SchedulerConfig",
    "search.apps.SearchConfig",
    "pricing.apps.PricingConfig",
    "payment.apps.PaymentConfig",
    "comms.apps.CommsConfig",
    "log.apps.LogConfig"
)
COMMON_APPS = (
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "django_filters",
    "graphene_django",
    "graphene_graphiql_explorer",
    "rest_framework.authtoken",
    "rest_framework",
    "corsheaders",
)
INSTALLED_APPS = API_APPS + COMMON_APPS

MIDDLEWARE = [
    "corsheaders.middleware.CorsMiddleware",
    "django.middleware.security.SecurityMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
    "django.middleware.common.BrokenLinkEmailsMiddleware",
    "django.middleware.common.CommonMiddleware",
]

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.debug",
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

# Database
# https://docs.djangoproject.com/en/2.2/ref/settings/#databases

DATABASES = env.DATABASES


# Password validation
# https://docs.djangoproject.com/en/2.2/ref/settings/#auth-password-validators

AUTH_PASSWORD_VALIDATORS = [
    {
        "NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.MinimumLengthValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.CommonPasswordValidator",
    },
    {
        "NAME": "django.contrib.auth.password_validation.NumericPasswordValidator",
    },
]

REST_FRAMEWORK = {
    "DEFAULT_AUTHENTICATION_CLASSES": [
        "rest_framework.authentication.TokenAuthentication",
    ],
    "TEST_REQUEST_DEFAULT_FORMAT": "json",
    "UNAUTHENTICATED_USER": None,
    # throttling
    "DEFAULT_THROTTLE_RATES": {
        "anon": "100/hour",
        "user": "1000/hour",
    }
}

ROOT_URLCONF = "mainframe.urls"
SECRET_KEY = env.SECRET_KEY
LANGUAGE_CODE = "en-us"
TIME_ZONE = "UTC"
USE_I18N = True
USE_L10N = True
USE_TZ = True
WSGI_APPLICATION = "mainframe.wsgi.application"


# Static files (CSS, JavaScript, Images)
# https://docs.djangoproject.com/en/2.2/howto/static-files/

STATIC_URL = "/static/"
STATIC_ROOT = os.path.join(BASE_DIR, "static/")

# CORS settings
CORS_ORIGIN_ALLOW_ALL = False
CORS_ORIGIN_WHITELIST = (
    "http://localhost:8000",
    "http://localhost:3000",
    "https://www.omoulearning.com",
    "https://omoulearning.com",
    "https://development.omoulearning.com",
    "https://www.development.omoulearning.com"
)

# GraphQL
GRAPHENE = {
    'SCHEMA': 'mainframe.schema.schema',
    'DJANGO_CHOICE_FIELD_ENUM_V3_NAMING': True,
    'MIDDLEWARE': [
        'graphql_jwt.middleware.JSONWebTokenMiddleware',
    ],
}
AUTHENTICATION_BACKENDS = [
    'graphql_jwt.backends.JSONWebTokenBackend',
    'django.contrib.auth.backends.ModelBackend',
]
SENDGRID_API_KEY = env.SENDGRID_API_KEY
TWILIO_ACCOUNT_SID = 'test'
TWILIO_AUTH_TOKEN = 'test'
BUSINESS_NAME = 'Stark Industries'
