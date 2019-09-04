#!/bin/sh


DATE=`date -Idate`
mDATE=`date`
echo $DATE

sh << EOF

cd /opt/rdms/py-email/

echo "$mDATE ... ing" >> rx-email.log
python recv_email.py sw64419 >> rx-email.log
echo "$mDATE ... ed" >> rx-email.log

EOF

#
