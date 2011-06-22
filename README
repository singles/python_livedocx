python_livedocx is wrapper for [LiveDocx](http://livedocx.com) service. It simplifies usage of theirs API.
It's distributed on MIT license.
Requirements: Python > 2.6.1, [SUDS](https://fedorahosted.org/suds/)

USAGE (download sample template from here http://www.phplivedocx.org/articles/brief-introduction-to-phplivedocx/):

    from livedocx import LiveDocx

    ld = LiveDocx()
    ld.login('username', 'password')
    ld.set_local_template('path/to/template.doc')

    ld.assign_value('software', 'python_livedocx')
    ld.assign_value('license', 'MIT')

    ld.create_document()

    data = ld.retrieve_document('PDF')

    file = open('software info.pdf', 'wb')
    file.write(data)
    file.close()

CHANGELOG:

===========================================
 v. 0.1
===========================================

1. Initial release