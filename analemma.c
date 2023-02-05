/* analemma.c -- code for computing analemmas for various      */
/*               orbital parameters                            */
/*                                                             */
/*     Notes: Computes and plots analemma curves.  Should      */
/*     compile OK under Unixes with -lm to link in the math    */
/*     library.  Needs xgraph to actually display the curves;  */
/*     use the -DHAVE_XGRAPH option if you have it.  If you    */
/*     don't, output goes to /tmp/analemma.xg for you to       */
/*     examine.                                                */
/*                                                             */
/*     First rotates a point westward for rotational effect,   */
/*     then rotates it eastward for orbital effect.  For       */
/*     further details, see the August 2002 Astronomical       */
/*     Games column.                                           */
/*                                                             */
/* Copyright (c) 2002 Brian Tung <brian@isi.edu>               */

#include <stdio.h>
#include <math.h>
#include <unistd.h>
#include <stdlib.h>

#ifndef PI
#define PI (3.1415926535)
#endif
#define DEGREES (PI/180.0)

extern char *optarg;

double ecc = 0.01671;     /* orbital eccentricity */
double lon = 1.347;       /* longitude of perihelion in radians */
double obliq = 0.4091;    /* obliquity in radians */

main(int argc, char **argv)
{
    FILE *xg;

    double theta;
    double t;
    double tau;
    double f;

    double x1, y1, z1;
    double x2, y2, z2;
    double x3, y3, z3;

    double eot;
    double dec;

    int eotMode = 0;
    int decMode = 0;

    char c;

    /* parse arguments */
    while ((c = getopt(argc, argv, "e:l:o:dqh")) >= 0) {
        switch (c) {
          case 'e':
            ecc = atof(optarg);
            break;
          case 'l':
            lon = atof(optarg)*DEGREES;
            break;
          case 'o':
            obliq = atof(optarg)*DEGREES;
            break;
          case 'd':
            decMode = 1;
            eotMode = 0;
            break;
          case 'q':
            decMode = 0;
            eotMode = 1;
            break;
          default:
            fprintf(stderr, "Usage: analemma [options]\n");
            fprintf(stderr, "    -e <ecc>    eccentricity ");
            fprintf(stderr, "(default value: %.5f)\n", ecc);
            fprintf(stderr, "    -l <lon>    longitude of perihelion in deg ");
            fprintf(stderr, "(default value: %.2f)\n", lon/DEGREES);
            fprintf(stderr, "    -o <obliq>  axial obliquity in deg ");
            fprintf(stderr, "(default value: %.2f)\n", obliq/DEGREES);
            fprintf(stderr, "    -d          plot declination instead\n");
            fprintf(stderr, "    -q          plot equation of time instead\n");
            fprintf(stderr, "    -h          print this page\n");
            exit(0);
        }
    }

    xg = fopen("/tmp/analemma.xg", "w");
    if (!xg)
        exit(-1);

    for (f = 0.0; f <= 1.0; f += 0.0001) {
        tau = 2.0*PI*f;
        /* first set theta to current longitude */
        theta = atan2(sqrt(1.0-ecc*ecc)*sin(tau), cos(tau)-ecc);

        /* first rotate clockwise in x-y plane by theta, corrected by lon */
        x1 = cos(theta-(lon-PI/2.0));
        y1 = sin(theta-(lon-PI/2.0));
        z1 = 0.0;

        /* secondly, rotate counter-clockwise in x-z plane by obliq */
        x2 = cos(obliq)*x1+sin(obliq)*z1;
        y2 = y1;
        z2 = -sin(obliq)*x1+cos(obliq)*z1;

        /* lastly, set t equal to real time from tau and
           rotate counter-clockwise by t, corrected by lon */
        t = tau-ecc*sin(tau);
        x3 = cos(t-(lon-PI/2.0))*x2+sin(t-(lon-PI/2.0))*y2;
        y3 = -sin(t-(lon-PI/2.0))*x2+cos(t-(lon-PI/2.0))*y2;
        z3 = z2;

        eot = -atan2(y3, x3)*4.0/DEGREES;
        dec = asin(z3)/DEGREES;

        /* print results in minutes early/late and degrees north/south */
        if (decMode)
            fprintf(xg, "%.9f %.9f\n", t/(2.0*PI), dec);
        else if (eotMode)
            fprintf(xg, "%.9f %.9f\n", t/(2.0*PI), eot);
        else
            fprintf(xg, "%.9f %.9f\n", eot, dec);
    }

    fclose(xg);

#ifdef HAVE_XGRAPH
    if (decMode)
        system("xgraph -0 Declination -x Time -y Dec /tmp/analemma.xg");
    else if (eotMode)
        system("xgraph -0 'Equation of Time' -x Time -y EOT /tmp/analemma.xg");
    else
        system("xgraph -0 Analemma -x EOT -y Dec /tmp/analemma.xg");

    unlink("/tmp/analemma.xg");
#endif

    exit(0);
}
