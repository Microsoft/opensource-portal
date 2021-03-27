//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

import sslify from 'express-sslify';

export default sslify.HTTPS({ trustAzureHeader: true });
